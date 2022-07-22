Attribute VB_Name = "Globals"
Option Explicit
' API declarations
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long
Public Const EM_SETCUEBANNER As Long = &H1501
Public Const VK_TAB = &H9
Public Const VK_LMENU = &HA4

' Application constants
Public Const BGCOLOR_INVALID As Long = 13421823
Public Const FGCOLOR_INVALID As Long = vbBlack
Public Const DATE_RE As String = "^([0-9][0-9])\.([0-9][0-9])\.([0-9][0-9][0-9][0-9])$"
Public Const DATE_HUMANREADABLE As String = "dd.mm.yyyy"
Public Const STATE_ONEWAY As Integer = 0
Public Const STATE_RETURN As Integer = 1
