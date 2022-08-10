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

Private Const DIGITS As String = "0123456789"
Private Const ALPHA As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
Private Const ALPHANUM As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
Private Const OPERATORS As String = "+-*/"
Private Const PARENS As String = "()"

Private Enum TokenizerState
    Initial
    Identifier
    Number
End Enum

Private m_Resolver As Resolver
Private m_Expression As String
Private m_ParsedExpression As Queue
Private m_ExternalRefs As Collection

Public Property Get Expression() As String
    Expression = m_Expression
End Property


Public Property Let Expression(ByVal Value As String)
    If Len(Value) = 0 Then
        Err.Raise 5, _
            Source:="Expression::Expression_Let", _
            Description:="Expression cannot be empty"
    End If
    m_Expression = Value
    Parse
End Property


Public Property Get Resolver() As Resolver
    Set Resolver = m_Resolver
End Property


Public Property Set Resolver(ByRef Value As Resolver)
    Set m_Resolver = Value
End Property


Public Function Evaluate() As Variant
    If Len(Expression) = 0 Then
        Err.Raise 5, _
            Source:="Expression::Evaluate", _
            Description:="Expression cannot be empty"
    End If
    If m_ParsedExpression Is Nothing Then
        Parse
    End If
    Dim evalStack As New Stack
    Dim expr As Queue, token As Variant
    Set expr = m_ParsedExpression.Clone()
    Do While expr.Count > 0
        token = expr.Dequeue()
        If InStr(OPERATORS, CStr(token)) = 0 Then
            evalStack.Push token
        Else
            Dim arg1 As Variant, arg2 As Variant
            If evalStack.Count < 2 Then
                Err.Raise vbObjectError + 203, _
                    Source:="Expression::Evaluate", _
                    Description:="Error in formula"
            End If
            arg2 = evalStack.Pop
            arg1 = evalStack.Pop
            If Not IsNumeric(arg1) Then
                arg1 = Resolver.Resolve(CStr(arg1))
            End If
            If Not IsNumeric(arg2) Then
                arg2 = Resolver.Resolve(CStr(arg2))
            End If
            Select Case token
                Case "+"
                    evalStack.Push arg1 + arg2
                Case "-"
                    evalStack.Push arg1 - arg2
                Case "*"
                    evalStack.Push arg1 * arg2
                Case "/"
                    evalStack.Push arg1 / arg2
            End Select
        End If
    Loop
    Evaluate = evalStack.Pop()
End Function


''' \brief Converts the formula into RPN form for evaluation using the
''' shunting-yard algorithm
''' \see https://en.wikipedia.org/wiki/Shunting_yard_algorithm
Private Sub Parse()
    Set m_ParsedExpression = Nothing
    Set m_ExternalRefs = New Collection
    Dim token As Variant
    Dim opStack As New Stack
    Dim outputQueue As New Queue
    Dim gotParen As Boolean
    For Each token In GetTokens()
        If Len(CStr(token)) = 1 And InStr(OPERATORS, CStr(token)) <> 0 Then
            Do
                If opStack.IsEmpty Then
                    Exit Do
                ElseIf CStr(opStack.Peek) <> "(" And _
                    GetPrecedence(CStr((opStack.Peek))) >= GetPrecedence(CStr(token)) Then
                    outputQueue.Enqueue opStack.Pop()
                Else
                    Exit Do
                End If
            Loop
            opStack.Push token
        ElseIf CStr(token) = "(" Then
            opStack.Push token
        ElseIf CStr(token) = ")" Then
            gotParen = False
            Do While opStack.Count > 0
                If CStr(opStack.Peek) = "(" Then
                    gotParen = True
                    Exit Do
                End If
                outputQueue.Enqueue opStack.Pop()
            Loop
            If Not gotParen Then
                Err.Raise vbObjectError + 202, _
                    Source:="Expression::Parse", _
                    Description:="Mismatched parenthesis in formula"
            End If
        Else
            outputQueue.Enqueue token
        End If
    Next
    ' Anything left on the stack?
    Do While opStack.Count > 0
        If CStr(opStack.Peek) = "(" Then
            Err.Raise vbObjectError + 202, _
                Source:="Expression::Parse", _
                Description:="Mismatched parenthesis in formula"
        End If
        outputQueue.Enqueue opStack.Pop()
    Loop
    Set m_ParsedExpression = outputQueue
End Sub


Private Function GetTokens() As Collection
    Dim state As TokenizerState
    Dim i As Integer, currentChar As String
    Dim tokenBuffer As String, advance As Boolean
    Dim result As New Collection
    
    state = Initial: tokenBuffer = "": i = 1
    Do While i <= Len(m_Expression)
        advance = True
        currentChar = Mid(m_Expression, i, 1)
        Select Case state
            Case Initial
                ' Eat whitespace
                If InStr(" " & vbTab & vbCrLf, currentChar) <> 0 Then
                    currentChar = ""
                ElseIf InStr(DIGITS, currentChar) <> 0 Then
                    state = Number
                ElseIf InStr(ALPHA, UCase(currentChar)) <> 0 Then
                    state = Identifier
                ' Handle operators as special case because they're only a
                ' single character long
                ElseIf InStr(OPERATORS, currentChar) <> 0 Or _
                    InStr(PARENS, currentChar) <> 0 _
                Then
                    result.Add currentChar
                    currentChar = ""
                Else
                    Err.Raise vbObjectError + 200, _
                    Source:="Expression::GetTokens", _
                    Description:=Strings.Format("Unexpected character ""{0}"" at postion {1}", currentChar, i)
                End If
            Case Number
                If currentChar = "." Then
                    If InStr(tokenBuffer, ".") <> 0 Then
                        Err.Raise vbObjectError + 201, _
                            Source:="Expression::GetTokens", _
                            Description:="Error in formula"
                    ElseIf InStr(DIGITS, currentChar) = 0 Then
                        ' End of token
                        result.Add tokenBuffer
                        tokenBuffer = ""
                        currentChar = ""
                        advance = False
                        state = Initial
                    End If
                End If
            Case Identifier
                If InStr(ALPHANUM, UCase(currentChar)) = 0 Then
                    ' End of token
                    tokenBuffer = UCase(tokenBuffer)
                    result.Add tokenBuffer
                    ' Add to external refs so the dependency graph can be calculated
                    ExternalReferences.Add tokenBuffer
                    tokenBuffer = ""
                    currentChar = ""
                    advance = False
                    state = Initial
                End If
        End Select
        tokenBuffer = tokenBuffer & currentChar
        If advance Then i = i + 1
    Loop
    If Len(tokenBuffer) > 0 Then
        result.Add tokenBuffer
    End If
    Set GetTokens = result
End Function


Public Function GetPrecedence(ByVal Operator As String) As Integer
    ' This is a hack! Will need to have a proper table if the language gets
    ' more complicated
    GetPrecedence = InStr(OPERATORS, Operator) \ 2
End Function


Public Property Get ExternalReferences() As Collection
    Set ExternalReferences = m_ExternalRefs
End Property
