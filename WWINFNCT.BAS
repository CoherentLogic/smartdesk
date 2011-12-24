Attribute VB_Name = "modW32StdFunctions"
'WWINFNCT.BAS
' WillisWare Standard Functions/Subroutines
' (Windows Version)

Public Const SWP_NOZORDER = &H4
Public Const FLAGS = 2 Or 1
Public Const HWND_TOPMOST = -1
Public Temp$
Public RunList(1 To 1000) As Double
Public RunListText(1 To 1000) As String
Public RunIndex As Integer


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public HelpData$
Public Sub AddXOSVar(VarName As String, VarValue As String)
    Call ProfileWrite("C:\XOS\USER\PROFILE", "Environment", VarName, VarValue)
End Sub


Public Sub MakeTopWindow(fnm As Form)
    Ret& = SetWindowPos(fnm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    
End Sub

Function nDirExists(DirName As String) As Integer
    Dim OldDir As String
    On Local Error Resume Next
    
    'Save the current directory
    OldDir = CurDir$

    'Attempt to change to the directory
    ChDir DirName

    'Check the error code
    If Err > 0 Then
        nDirExists = False
    Else
        nDirExists = True
    End If
    ChDir "\" + OldDir
End Function

Function nExists(FileName As String) As Integer
    Dim nExistsFile As Integer
    On Local Error Resume Next

    'Attempt to open the file
    nExistsFile = FreeFile
    Open FileName For Input As #nExistsFile

    'Check the error code
    If Err > 0 Then
        nExists = False
    Else
        nExists = True
    End If

    'Close the file
    Close #nExistsFile
End Function

Function nFigurePercent(Given As Double, Max As Double) As Double
    On Local Error Resume Next
    Dim GivenMod1 As Double
    'Percentage Algorithm
    GivenMod1 = 100 * Given
    nFigurePercent = Int(GivenMod1 / Max)
End Function

Function nReadProfileInt(a$, b$, c$) As Integer
    nReadProfileInt = Val(ProfileRead$(a$, b$, c$))
End Function


Function szGetRootDir(Path As String, File As String) As String
    If Right$(Path, 1) = "\" Then
        szGetRootDir = Path + File
    Else
        szGetRootDir = Path + "\" + File
    End If
End Function


Function szKillChar(InText As String, TrimChar As String, ReplaceWith As String) As String

    For Iteration = 1 To Len(InText)
        'Get one character from the text
        Temp$ = Mid$(InText, Iteration, 1)

        'If the character is the same as
        'TrimChar, exclude it. Otherwise,
        'include it in Whole$
        If Not UCase$(Temp$) = UCase$(TrimChar) Then
            Whole$ = Whole$ + Temp$
        Else
            Whole$ = Whole$ + ReplaceWith
        End If
    Next Iteration

    szKillChar = Whole$
End Function

Function szStringLeft(InText As String, CharFilter As String) As String

    If InStr(InText, CharFilter) > 0 Then
        nCharPos% = InStr(InText, CharFilter)
        Temp$ = Mid$(InText, 1, nCharPos%)
        szStringLeft = Temp$
    Else
        szStringLeft = InText
    End If

End Function


Function szStringRight(InText As String, CharFilter As String) As String

    If Not InStr(InText, CharFilter) > 0 Then
        szStringRight = InText
        Exit Function
    End If

    'Get the characters
    For Iteration% = Len(InText) To 1 Step -1

        Temp$ = Mid$(InText, Iteration%, 1)
        
        If Temp$ = CharFilter Then
            For Iteration2% = Len(Whole$) To 1 Step -1

                Temp$ = Mid$(Whole$, Iteration2%, 1)
                Result$ = Result$ + Temp$

            Next Iteration2%

            szStringRight = Result$
            Exit Function
        Else
            Whole$ = Whole$ + Temp$
        End If
    Next Iteration%

    'Reverse the string before returning the result

End Function

Function szTrimText(InText As String) As String
    szTrimText = LTrim$(RTrim$(InText))
End Function

Public Function szXOSVar(VarName As String) As String
    szXOSVar = ProfileRead("C:\XOS\USER\PROFILE", "Environment", VarName)
    
End Function


