Attribute VB_Name = "modStdDatabase"
Public hFile As Integer
Public HelpFile As String
Public Sub CloseFile()
    Close #hFile
End Sub


Public Function szExtractDataSection(InData As String, SectionName As String) As String

Dim nIteration As Integer
Dim hFile As Integer
Dim Rawline As String
Dim nInSection As Boolean
Dim szBuildup As String
Dim FullMarker As String
hFile = FreeFile

    Open "C:\XOS\TMP\SZDATA" For Output As #hFile
    
    Print #hFile, InData
    Close #hFile
    
    hFile = FreeFile
    
    Open "C:\XOS\TMP\SZDATA" For Input As #hFile
    
    Do While Not EOF(hFile)
ReLoop:
        
        Line Input #hFile, Rawline
        
        If Left(Rawline, 1) = "$" And InStr(Rawline, SectionName) Then
            nInSection = True
            FullMarker = Rawline
            GoTo ReLoop
        End If
        
        If nInSection = True And Left$(Rawline, 4) = "*END" Then
            nInSection = False
            szExtractDataSection = szBuildup
            Close #hFile
            Exit Function
        End If
        
        If nInSection = True Then
            szBuildup = szBuildup & Rawline & vbCrLf
        End If
    
    Loop
    
    Close #hFile
    szExtractDataSection = szBuildup
    
    
    
End Function


Public Function szReadFromFile(szFilename As String) As String
On Error Resume Next

hFile = FreeFile
    
    Open szFilename For Input As #hFile
    
    szReadFromFile = Input$(LOF(hFile), hFile)
    Close #hFile
    
End Function


