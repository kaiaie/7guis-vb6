VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Expression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum TokenizerState
    Initial
    Identifier
    Number
    StringLiteral
End Enum

Private m_Expression As String
Private m_ParsedExpression As Object

Public Property Get Expression() As String
    Expression = m_Expression
End Property


Public Property Let Expression(ByVal Value As String)
    If Len(Value) = 0 Then
        Err.Raise 5, Description:="Expression cannot be empty"
    End If
    m_Expression = Value
    Parse
End Property


Public Function Evaluate() As Variant
    Evaluate = vbNull
End Function


' Creates the parse tree for the formula
Private Sub Parse()
    Set m_ParsedExpression = Nothing
    Dim token As String
    For Each token In GetTokens()
    Next
End Sub


Private Function GetTokens() As Collection
    Const DIGITS As String = "0123456789"
    Const ALPHA As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
    Const ALPHANUM As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
    Const OPERATORS As String = "+-*/"
    Dim state As TokenizerState
    Dim i As Integer, currentChar As String
    Dim tokenBuffer As String
    Dim result As New Collection
    
    state = Initial: tokenBuffer = ""
    For i = 1 To Len(m_Expression)
        currentChar = Mid(m_Expression, i, 1)
        Select Case state
            Case Initial
                ' Eat whitespace
                If InStr(" " & vbTab & vbCrLf, currentChar) <> 0 Then
                    currentChar = ""
                ElseIf InStr(DIGITS, currentChar) <> 0 Then
                    state = Number
                ElseIf currentChar = """" Then
                    state = StringLiteral
                ElseIf InStr(ALPHA, currentChar) <> 0 Then
                    state = Identifier
                ' Handle operators as special case
                ElseIf InStr(OPERATORS, currentChar) <> 0 Then
                    result.Add currentChar
                    currentChar = ""
                End If
            Case Number
                If currentChar = "." Then
                    If InStr(tokenBuffer, ".") <> 0 Then
                        Err.Raise vbObjectError + 201, Description:="Error in formula"
                    ElseIf InStr(DIGITS, currentChar) = 0 Then
                        ' End of token
                        result.Add tokenBuffer
                        tokenBuffer = currentChar
                        currentChar = ""
                        state = Initial
                    End If
                End If
            Case Identifier
                If InStr(ALPHANUM, currentChar) = 0 Then
                    ' End of token
                    result.Add tokenBuffer
                    tokenBuffer = currentChar
                    currentChar = ""
                    state = Initial
                End If
            Case StringLiteral
                If currentChar = """" Then
                    ' End of token
                    result.Add tokenBuffer & currentChar
                    tokenBuffer = ""
                    currentChar = ""
                    state = Initial
                End If
        End Select
        tokenBuffer = tokenBuffer & currentChar
    Next
    
    Set GetTokens = result
End Function
