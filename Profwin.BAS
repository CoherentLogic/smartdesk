Attribute VB_Name = "mouil"
' Notes:
' Add '$INCLUDE: 'PROFILE.BI' to the calling module.
' Load PROFILE.BAS into your project list.
'
' Example of use:
'    Last$ = ProfileRead(IniFile$, IniSection$, IniKey$)
'    ProfileWrite IniFile$, IniSection$, IniKey$, ProfileStr$

Public HeaderData(100) As String
Public NumHeaders As Integer

Public Sub GetProfileHeaders(ProfDoc As String)
    u = FreeFile
    Open ProfDoc For Input As u
    
    Do While Not EOF(u)
        Line Input #u, Rawline$
        
        If Left$(Rawline$, 1) = "[" Then
            p$ = Mid$(Rawline$, 2)
            NextHeader$ = Mid$(p$, 1, InStr(p$, "]") - 1)
            nInd = nInd + 1
            HeaderData(nInd) = NextHeader$
           
        End If
    Loop
    NumHeaders = nInd
    Close u
End Sub

'$INCLUDE: 'CONSTANT.BI'
'$INCLUDE: 'PROFILE.BI'

Function ProfileRead$(IniFile$, IniSection$, IniKey$)
    TmpSection$ = "[" + IniSection$ + "]"
    On Local Error GoTo BadIniReadFile
    FileNum% = FreeFile
    Open IniFile$ For Input As #FileNum%
    Do Until EOF(FileNum%)
        Line Input #FileNum%, a$
        If Left$(a$, 1) = "[" Then
            If InStr(a$, TmpSection$) = 1 Then
                SectionFlag% = True
            Else
                SectionFlag% = False
            End If
        End If
        If SectionFlag% = True Then
            If InStr(a$, IniKey$) = 1 Then
                a$ = Mid$(a$, InStr(a$, "=") + 1)
                ProfileRead$ = LTrim$(RTrim$(a$))
                Exit Do
            End If
        End If
    Loop
            
AllDoneRead:
    Close #FileNum%
    Exit Function

BadIniReadFile:
    Resume AllDoneRead
End Function

Sub ProfileWrite(IniFile$, IniSection$, IniKey$, ProfileStr$)
    ReDim p$(1)

    TmpSection$ = "[" + IniSection$ + "]"
    FileNum% = FreeFile

    On Local Error GoTo NoIniWriteFile
    Open IniFile$ For Input As #FileNum%
    Do Until EOF(FileNum%)
        n% = n% + 1
        ReDim Preserve p$(n%)
        Line Input #FileNum%, p$(n%)
    Loop
    Close #FileNum%

AddToIniFile:
    On Local Error GoTo BadIniWriteFile
    For i% = 1 To n%
        If Left$(p$(i%), 1) = "[" Then
            If InStr(p$(i%), TmpSection$) = 1 Then
                SectionFlag% = True
                SectionFound% = i%
            Else
                SectionFlag% = False
            End If
        End If
        If SectionFlag% = True Then
            If InStr(p$(i%), IniKey$) = 1 Then
                p$(i%) = IniKey$ + "=" + ProfileStr$
                KeyFound% = True
                Exit For
            End If
        End If
    Next i%

    If SectionFound% = False Then
        ReDim Preserve p$(n% + 3)
        p$(n% + 2) = "[" + IniSection$ + "]"
        p$(n% + 3) = IniKey$ + "=" + ProfileStr$
        n% = n% + 3
    ElseIf KeyFound% = False Then
        NewLine% = SectionFound% + 1
        ReDim Preserve p$(n% + 1)
        For i% = n% To NewLine% Step -1
            p$(i% + 1) = p$(i%)
        Next i%
        p$(NewLine%) = IniKey$ + "=" + ProfileStr$
        n% = n% + 1
    End If

    Open IniFile$ For Output As #FileNum%
    For i% = 1 To n%
        Print #FileNum%, p$(i%)
    Next i%
    Close #FileNum%

AllDoneWrite:
    Close #FileNum%
    Exit Sub

NoIniWriteFile:
    Open IniFile$ For Output As #FileNum%
    Print #FileNum%, Space$(1)
    Close #FileNum%
    Resume AddToIniFile

BadIniWriteFile:
    Resume AllDoneWrite
End Sub

