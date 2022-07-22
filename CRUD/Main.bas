Attribute VB_Name = "MainModule"
Option Explicit

Public Sub Main()
    Dim testPeople As People
    Set testPeople = CreateSomePeople()
    Dim frm As New CRUDForm
    Set frm.DataSource = testPeople
    frm.Show
End Sub


Public Function CreateSomePeople() As People
    Dim result As New People
    Dim p As Person
    
    Set p = New Person
    With p
        .Name = "Hans"
        .Surname = "Emil"
    End With
    result.Add p
    Set p = New Person
    With p
        .Name = "Max"
        .Surname = "Mustermann"
    End With
    result.Add p
    Set p = New Person
    With p
        .Name = "Roman"
        .Surname = "Tisch"
    End With
    result.Add p
    
    Set CreateSomePeople = result
End Function
