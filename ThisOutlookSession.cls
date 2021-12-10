
' Variablen die evtl. in Sub angepasst werden müssen:
Dim sSubject As String
Dim beschreibung As String
Dim sLocation As String
Dim iDuration As Integer
Dim iXth As Integer
Dim iAmount As Integer

Sub Serientermin()
Dim oDate As Date
Dim oDateTmp As Date
Dim iMonth
Dim iDay As Integer
Dim iYear

Dim iBeginn As String
Dim iEnde As String
Dim iXth
Dim beginn As Integer
Dim monat As String
Dim error

'On Error GoTo FehlerVerarbeitung
' Initialisierung - 1. des derzeitigen Monats
iDay = 1

'Userform leeren
UserForm1.titelBox = ""
UserForm1.TextBox1 = ""

UserForm1.Show

iXth = UserForm1.ComboTag.Value
iMonth = UserForm1.ComboMonat.Value

If iMonth = "Januar" Then
    iMonth = 1
ElseIf iMonth = "Februar" Then
    iMonth = 2
ElseIf iMonth = "März" Then
    iMonth = 3
ElseIf iMonth = "April" Then
    iMonth = 4
ElseIf iMonth = "Mai" Then
    iMonth = 5
ElseIf iMonth = "Juni" Then
    iMonth = 6
ElseIf iMonth = "Juli" Then
    iMonth = 7
ElseIf iMonth = "August" Then
    iMonth = 8
ElseIf iMonth = "September" Then
    iMonth = 9
ElseIf iMonth = "Oktober" Then
    iMonth = 10
ElseIf iMonth = "November" Then
    iMonth = 11
ElseIf iMonth = "Dezember" Then
    iMonth = 12
End If

If (iMonth = 4 Or iMonth = 6 Or iMonth = 9 Or iMonth = 11 And iXth = 31) Then
    iXth = 30
End If


iYear = UserForm1.ComboJahr.Value

oDate = CDate(iXth & "." & iMonth & "." & iYear)

sSubject = UserForm1.titelBox.Value
beschreibung = UserForm1.TextBox1.Value
iBeginn = UserForm1.ComboBeginn.Value
iEnde = UserForm1.ComboEnde.Value
iAmount = UserForm1.AnzahlTermine.Value


beginn = CInt(Mid(iBeginn, 1, 2))

' Variablen-Ini
Dim i As Integer
Dim myItem As AppointmentItem
Dim myOlApp As Outlook.Application
Dim objRecPatt As RecurrencePattern

Set myOlApp = GetObject(, "Outlook.Application")


For i = 0 To iAmount - 1
Set myItem = myOlApp.CreateItem(olAppointmentItem)

myItem.MeetingStatus = olMeeting
myItem.Subject = sSubject
myItem.Body = beschreibung

oDateTmp = DateAdd("m", i, oDate)
If ((Mid(oDateTmp, 4, 2) = 1 Or Mid(oDateTmp, 4, 2) = 3 Or Mid(oDateTmp, 4, 2) = 5 Or Mid(oDateTmp, 4, 2) = 7 _
Or Mid(oDateTmp, 4, 2) = 8 Or Mid(oDateTmp, 4, 2) = 10 Or Mid(oDateTmp, 4, 2) = 12) And (iXth = 30 Or iXth = 28)) Then
    iXth = 31
ElseIf ((Mid(oDateTmp, 4, 2) = 2 And iXth = 30) Or (Mid(oDateTmp, 4, 2) = 2 And iXth = 31)) Then
    iXth = 28
ElseIf ((Mid(oDateTmp, 4, 2) = 4 Or Mid(oDateTmp, 4, 2) = 6 Or Mid(oDateTmp, 4, 2) = 9 Or Mid(oDateTmp, 4, 2) = 11) _
And (iXth = 31 Or iXth = 28)) Then
    iXth = 30
End If

If i > 0 Then
    If iMonth = 12 Then
        iMonth = 1
        iYear = iYear + 1
    Else
        iMonth = iMonth + 1
    End If
End If

oDateTmp = CDate(iXth & "." & iMonth & "." & iYear)

getXthWorkingDay oDateTmp, iXth

myItem.Start = CStr(oDateTmp) & " " & iBeginn & ":00 AM"
myItem.End = CStr(oDateTmp) & " " & iEnde & ":00 AM"

myItem.Send
Next i

'FehlerVerarbeitung:
' error = MsgBox("Fehler", vbOKOnly, "Error")
'Exit Sub

UserForm2.Show


End Sub

Private Sub getXthWorkingDay(oDate, iX)
Dim a
Dim oDateTmp
Dim oDayOfWeek

    oDateTmp = oDate
    oDayOfWeek = Format(oDateTmp, "dddd")
    datumTag = CInt(Mid(oDateTmp, 1, 2))
    
    
'Wenn Termin Kalendertagsbezogen ist, wird der Termin vorgezogen
 If (UserForm1.CheckBox1.Value = True) Then
  If oDayOfWeek = "Samstag" Then
            oDateTmp = oDateTmp - 1 'Springt auf den Freitag
    ElseIf oDayOfWeek = "Sonntag" Then
            oDateTmp = oDateTmp - 2 'Springt auf den Freitag
    Else: oDateTmp = oDateTmp ' Bleibt
        End If
 End If
 
 
'Wenn Termin Arbeitstagsbezogen ist und am letzten AT des Monats ist,
'wird der Termin vorgezogen
If (UserForm1.CheckBox1.Value = False) Then

    If (datumTag = 30 Or datumTag = 31) Then
        If oDayOfWeek = "Samstag" Then
                oDateTmp = oDateTmp - 1 'Springt auf den Freitag
        ElseIf oDayOfWeek = "Sonntag" Then
                oDateTmp = oDateTmp - 2 'Springt auf den Freitag
        Else: oDateTmp = oDateTmp ' Bleibt
            End If
    End If
    
    'Wenn Termin Arbeitstagsbezogen ist und am ersten AT des Monats ist,
    'wird der Termin vorgezogen
    If Not (datumTag = 30 Or datumTag = 31) Then
        If oDayOfWeek = "Samstag" Then
                oDateTmp = oDateTmp + 2 'Springt auf den Montag
        ElseIf oDayOfWeek = "Sonntag" Then
                oDateTmp = oDateTmp + 1 'Springt auf den Montag
        Else: oDateTmp = oDateTmp ' Bleibt
            End If
    End If
End If

oDate = oDateTmp

End Sub


