VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Termin einstellen"
   ClientHeight    =   8310
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7035
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    
    Me.Hide
    
End Sub


Private Sub UserForm_Initialize()
With ComboBeginn
    .AddItem "00:00"
    .AddItem "00:30"
    .AddItem "01:00"
    .AddItem "01:30"
    .AddItem "02:00"
    .AddItem "02:30"
    .AddItem "03:00"
    .AddItem "03:30"
    .AddItem "04:00"
    .AddItem "04:30"
    .AddItem "05:00"
    .AddItem "05:30"
    .AddItem "06:00"
    .AddItem "06:30"
    .AddItem "07:00"
    .AddItem "07:30"
    .AddItem "08:00"
    .AddItem "08:30"
    .AddItem "09:00"
    .AddItem "09:30"
    .AddItem "10:00"
    .AddItem "10:30"
    .AddItem "11:00"
    .AddItem "11:30"
    .AddItem "12:00"
    .AddItem "12:30"
    .AddItem "13:00"
    .AddItem "13:30"
    .AddItem "14:00"
    .AddItem "14:30"
    .AddItem "15:00"
    .AddItem "15:30"
    .AddItem "16:00"
    .AddItem "16:30"
    .AddItem "17:00"
    .AddItem "17:30"
    .AddItem "18:00"
    .AddItem "18:30"
    .AddItem "19:00"
    .AddItem "19:30"
    .AddItem "20:00"
    .AddItem "20:30"
    .AddItem "21:00"
    .AddItem "21:30"
    .AddItem "22:00"
    .AddItem "22:30"
    .AddItem "23:00"
    .AddItem "23:30"
    
    .ListIndex = 16
End With

With ComboEnde
    .AddItem "00:00"
    .AddItem "00:30"
    .AddItem "01:00"
    .AddItem "01:30"
    .AddItem "02:00"
    .AddItem "02:30"
    .AddItem "03:00"
    .AddItem "03:30"
    .AddItem "04:00"
    .AddItem "04:30"
    .AddItem "05:00"
    .AddItem "05:30"
    .AddItem "06:00"
    .AddItem "06:30"
    .AddItem "07:00"
    .AddItem "07:30"
    .AddItem "08:00"
    .AddItem "08:30"
    .AddItem "09:00"
    .AddItem "09:30"
    .AddItem "10:00"
    .AddItem "10:30"
    .AddItem "11:00"
    .AddItem "11:30"
    .AddItem "12:00"
    .AddItem "12:30"
    .AddItem "13:00"
    .AddItem "13:30"
    .AddItem "14:00"
    .AddItem "14:30"
    .AddItem "15:00"
    .AddItem "15:30"
    .AddItem "16:00"
    .AddItem "16:30"
    .AddItem "17:00"
    .AddItem "17:30"
    .AddItem "18:00"
    .AddItem "18:30"
    .AddItem "19:00"
    .AddItem "19:30"
    .AddItem "20:00"
    .AddItem "20:30"
    .AddItem "21:00"
    .AddItem "21:30"
    .AddItem "22:00"
    .AddItem "22:30"
    .AddItem "23:00"
    .AddItem "23:30"
    
    .ListIndex = 17
End With

With ComboTag
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
    .AddItem "11"
    .AddItem "12"
    .AddItem "13"
    .AddItem "14"
    .AddItem "15"
    .AddItem "16"
    .AddItem "17"
    .AddItem "18"
    .AddItem "19"
    .AddItem "20"
    .AddItem "21"
    .AddItem "22"
    .AddItem "23"
    .AddItem "24"
    .AddItem "25"
    .AddItem "26"
    .AddItem "27"
    .AddItem "28"
    .AddItem "29"
    .AddItem "30"
    .AddItem "31"
    
    .ListIndex = 0
    
End With

With ComboMonat

    .AddItem "Januar"
    .AddItem "Februar"
    .AddItem "M�rz"
    .AddItem "April"
    .AddItem "Mai"
    .AddItem "Juni"
    .AddItem "Juli"
    .AddItem "August"
    .AddItem "September"
    .AddItem "Oktober"
    .AddItem "November"
    .AddItem "Dezember"
    
    .ListIndex = 0
    
End With


With ComboJahr
    .AddItem Year(Now)
    .AddItem Year(Now) + 1
    .AddItem Year(Now) + 2
    .AddItem Year(Now) + 3
    .AddItem Year(Now) + 4
    
    .ListIndex = 0
    
End With


titel = titelBox.Value
anzahl = AnzahlTermine.Value

End Sub
