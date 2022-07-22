VERSION 5.00
Begin VB.Form BookingForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Flight"
   ClientHeight    =   1680
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton B 
      Caption         =   "&Book"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2412
   End
   Begin VB.TextBox T2 
      Enabled         =   0   'False
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2412
   End
   Begin VB.TextBox T1 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2412
   End
   Begin VB.ComboBox C 
      Height          =   288
      ItemData        =   "Booking.frx":0000
      Left            =   120
      List            =   "Booking.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2412
   End
End
Attribute VB_Name = "BookingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Get State() As Integer
    State = C.ListIndex
End Property

Public Property Get DepartureDate() As Date
    Dim B As Boolean, dt As Date
    B = IsValidDate(T1.Text, dt)
    If B Then
        DepartureDate = dt
        Exit Property
    End If
    ' NOTE: It should not be possible to reach this
    Err.Raise vbObjectError + 300, _
        Description:="A valid departure date has not been entered"
End Property

Public Property Get ReturnDate() As Date
    Dim B As Boolean, dt As Date
    B = IsValidDate(T2.Text, dt)
    If B Then
        ReturnDate = dt
        Exit Property
    End If
    ' NOTE: It should not be possible to reach this
    Err.Raise vbObjectError + 301, _
        Description:="A valid return date has not been entered"
End Property


Private Sub B_Click()
    ShowBookingMessage
End Sub

Private Sub C_Click()
    UpdateForm
End Sub


Private Sub Form_Load()
    C.ListIndex = 0
    ' Note: this doesn't appear to work! It's possible the text must be coerced to Unicode first
    SendMessage T1.hwnd, EM_SETCUEBANNER, 0&, DATE_HUMANREADABLE
    SendMessage T2.hwnd, EM_SETCUEBANNER, 0&, DATE_HUMANREADABLE
End Sub


'* \brief Sets the enabled and validation state of the form's controls
Private Sub UpdateForm()
    Dim t1IsValidDate As Boolean, t2IsValidDate As Boolean
    Dim d1 As Date, d2 As Date
    
    t1IsValidDate = IsValidDate(T1.Text, d1)
    t2IsValidDate = IsValidDate(T2.Text, d2)
    SetTextBoxValidation T1, t1IsValidDate
    SetTextBoxValidation T2, t2IsValidDate
    T2.Enabled = (State = STATE_RETURN)
    B.Enabled = (State = STATE_ONEWAY And t1IsValidDate) Or _
        (State = STATE_RETURN And t1IsValidDate And t2IsValidDate And d2 > d1)
End Sub


'* \brief Sets a text box's colours if it contains invalid input
Private Sub SetTextBoxValidation(ByRef t As textBox, ByVal isValid As Boolean)
    If Len(t.Text) > 0 And Not isValid Then
        t.BackColor = Globals.BGCOLOR_INVALID
        t.ForeColor = Globals.FGCOLOR_INVALID
    Else
        t.BackColor = vbWindowBackground
        t.ForeColor = vbWindowText
    End If
End Sub


Private Function IsValidDate(ByVal s As String, ByRef dt As Date) As Boolean
    Static rx As VBScript_RegExp_55.RegExp
    If rx Is Nothing Then
        Set rx = New RegExp
        With rx
            .Pattern = Globals.DATE_RE
            .Global = True
            .IgnoreCase = True
        End With
    End If
    Dim matches As VBScript_RegExp_55.MatchCollection
    
    Set matches = rx.Execute(s)
    If matches.Count = 0 Then
        IsValidDate = False
        Exit Function
    End If
    ' Parse and return date
    Dim d As Integer: d = CInt(matches(0).SubMatches(0))
    Dim m As Integer: m = CInt(matches(0).SubMatches(1))
    Dim y As Integer: y = CInt(matches(0).SubMatches(2))
    dt = DateSerial(y, m, d)
    IsValidDate = True
End Function


Private Sub T1_Change()
    UpdateForm
End Sub


Private Sub T1_GotFocus()
    SelectTextOnKeyboadFocus T1
End Sub

Private Sub T2_Change()
    UpdateForm
End Sub


Private Sub ShowBookingMessage()
    Dim message As String
    
    If State = STATE_ONEWAY Then
        message = Strings.Format _
        ( _
            "You have booked a one-way flight departing on {0:dd/mm/yyyy}.", _
            DepartureDate _
        )
    Else
        message = Strings.Format _
        ( _
            "You have booked a return flight departing on {0:dd/mm/yyyy} and returning on {1:dd/mm/yyyy}.", _
            DepartureDate, ReturnDate _
        )
    End If
    MsgBox Prompt:=message, Buttons:=vbOKOnly Or vbInformation, Title:=Me.Caption
End Sub


'* \brief Selects the text in the control if the focus was moved there by
'* pressing Tab or using a Alt+character key, to be consistent with
'* normal dialog box behaviour
Private Sub SelectTextOnKeyboadFocus(ByRef textBox As textBox)
    Dim tabState As Integer: tabState = GetKeyState(VK_TAB)
    Dim altState As Integer: altState = GetKeyState(VK_LMENU)
    If tabState <> 0 Or altState <> 0 Then
        textBox.SelStart = 0
        textBox.SelLength = Len(textBox.Text)
    End If
End Sub

Private Sub T2_GotFocus()
    SelectTextOnKeyboadFocus T2
End Sub
