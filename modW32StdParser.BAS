Attribute VB_Name = "modW32StdParser"
'JWSD language parser (for interpreters)
' Version 1.00   (12/3/96)
'
' Specifications:
'   10 arguments/statement  (adjustable)
'   2 data types (string/numeric)
'   Data type identifiers are the leading characters of variable name
'   The type identifiers are user-definable 1-character strings
'   Contains functions for typical math operations (nAdd(), nSubtract(),
'       nMultiply(), nDivide(), nRemainder(), nAverage())
'   Does not as of yet support arrays of any type
'   Loads all arguments into a Visual Basic (r) array
'   Uses \xos\user\profile for variables
'
' Statement types:
'   1=Function call
'   2=String variable assignment
'   3=Numeric variable assignment
'
' Variable types:
'   1=String
'   2=Numeric
'
' Argument types:
'   1=Literal
'   2=Variable


Public LineArgs(1 To 10) As String
Public ArgSeperator As String
Public AssignmentOperator As String
Public InArgCharacter As String
Public StringIdent As String
Public NumericIdent As String
Public LiteralIdent As String
Public Function nArgType(Argument As String) As Integer
    If Left$(Argument, 1) = StringIdent Or Left$(Argument, 1) = NumericIdent Then
        nArgType = 2
    ElseIf Left$(Argument, 1) = LiteralIdent Then
        nArgType = 1
    End If
End Function


Public Function nReadVar(VarName As String) As Double
    nReadVar = Val(szReadVar(VarName))
End Function


Public Function nStatementType(Rawline As String) As Integer
    'If Left$(Rawline, 1) = StringIdent Then
        
End Function

Public Function nVarType(VarName As String) As Integer

End Function

Public Sub SetupParser(ArgSep As String, AOp As String, InArgChar As String)
    InArgCharacter = InArgChar
    ArgSeperator = ArgSep
    AssignmentOperator = AOp
End Sub

Public Sub SetVar(VarName As String, VarValue As String)
    Call ProfileWrite("C:\XOS\USER\PROFILE", "Environment", VarName, VarValue)
End Sub


Public Function szParseStatement(Rawline As String) As String
    
    
End Function

Public Function szReadVar(VarName As String) As String
    ReadVar = ProfileRead("C:\XOS\USER\PROFILE", "Environment", "VarName")
End Function


Public Sub ParseLine(Rawline As String)
          
    Dim Iteration As Integer
    Dim nArgIndex As Integer
    Dim WholeArg As String
    Dim Temp As String
          
    'Read each character one at a time
    For Iteration = 1 To Len(Rawline)
        Temp = Mid$(Rawline, Iteration, 1)
        
        If Temp = ArgSeperator Then
            nArgIndex = nArgIndex + 1
            LineArgs(nArgIndex) = WholeArg
        Else
            WholeArg = WholeArg & Temp
        End If
        
    Next Iteration

End Sub


