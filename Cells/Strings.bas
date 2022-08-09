Attribute VB_Name = "Strings"
Option Explicit

'* \brief Interpolates values into the specified format string, similar to .NET's
'* String.Format
'* \param formatString  The format string into which the other parameters will
'*                      be interpolated
'* \param params        The value(s) to interpolate
Public Function Format _
( _
    ByVal formatString As String, _
    ParamArray params() As Variant _
) As String
    Const initialState = 1
    Const startOfPlaceholderState = 2
    Const placeholderIndexState = 3
    Const paramFormatState = 4
    Const paramLengthInitialState = 5
    Const paramLengthState = 6
    Const replacePlaceholderState = 7

    Dim result As String
    Dim state As Integer: state = 1
    Dim pos As Integer: pos = 1
    Dim paramIdx As String
    Dim paramFormat As String
    Dim lengthBuffer As String
    
    Do
        If pos > Len(formatString) Then
            Exit Do
        End If
        Dim char As String
        char = Mid(formatString, pos, 1)
        ' Initial state
        If state = initialState Then
            If char = "{" Then
                state = startOfPlaceholderState
            Else
                result = result & char
            End If
            pos = pos + 1
        ' Possible start of placeholder
        ElseIf state = startOfPlaceholderState Then
            ' Two consecutive curly brackets is an escape
            If char = "{" Then
                result = result & char
                state = initialState
            Else
                If char < "0" Or char > "9" Then
                    Err.Raise vbObjectError + 1500, _
                        Description:="Bad character in placeholder"
                End If
                paramIdx = char
                state = placeholderIndexState
            End If
            pos = pos + 1
        ' Remaining characters of placeholder
        ElseIf state = placeholderIndexState Then
            If char = "}" Then
                ' End of placeholder
                state = replacePlaceholderState
            ElseIf char = "," Then
                ' Parameter length string (optional)
                lengthBuffer = ""
                state = paramLengthInitialState
                pos = pos + 1
            ElseIf char = ":" Then
                ' Parameter format string
                paramFormat = ""
                state = paramFormatState
                pos = pos + 1
            Else
                If char < "0" Or char > "9" Then
                    Err.Raise vbObjectError + 1500, _
                        Description:="Bad character in placeholder"
                End If
                paramIdx = paramIdx & char
                pos = pos + 1
            End If
        ElseIf state = paramFormatState Then
            If char = "}" Then
                ' End of placeholder
                state = replacePlaceholderState
            Else
                paramFormat = paramFormat & char
                pos = pos + 1
            End If
        ElseIf state = paramLengthInitialState Then
            If InStr("0123456789-", char) = 0 Then
                Err.Raise vbObjectError + 1501, _
                    Description:="Bad character in length value"
            End If
            lengthBuffer = lengthBuffer & char
            state = paramLengthState
            pos = pos + 1
        ElseIf state = paramLengthState Then
            If char = "}" Then
                ' End of placeholder
                state = replacePlaceholderState
            ElseIf char = ":" Then
                ' Parameter format string
                paramFormat = ""
                state = paramFormatState
                pos = pos + 1
            Else
                If char < "0" Or char > "9" Then
                    Err.Raise vbObjectError + 1501, _
                        Description:="Bad character in length value"
                End If
                lengthBuffer = lengthBuffer & char
                pos = pos + 1
            End If
        ElseIf state = replacePlaceholderState Then
            Dim idx As Integer: idx = CInt(paramIdx)
            If idx < LBound(params) Or idx > UBound(params) Then
                Err.Raise vbObjectError + 1502, _
                    Description:="Index out of range"
            End If
            Dim placeholder As String
            If Len(paramFormat) = 0 Then
                placeholder = CStr(params(idx))
            Else
                placeholder = VBA.Format(params(idx), paramFormat)
            End If
            Dim padded As String
            If Len(lengthBuffer) = 0 Then
                ' No length specified
                padded = placeholder
            Else
                Dim paddingSize As Integer
                paddingSize = CInt(lengthBuffer)
                If paddingSize = 0 Then
                    Err.Raise vbObjectError + 1503, _
                        Description:="Invalid parameter length, must be non-zero number"
                End If
                padded = Space(Abs(paddingSize))
                If paddingSize < 1 Then
                    ' Left align
                    LSet padded = placeholder
                Else
                    ' Right align
                    RSet padded = placeholder
                End If
            End If
            result = result & padded
            pos = pos + 1
            state = initialState
        End If
    Loop
    If state <> initialState Then
        Err.Raise vbObjectError + 1504, _
            Description:="Unclosed format placeholder"
    End If
    
    Format = result
End Function


