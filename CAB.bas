Attribute VB_Name = "CAB"
Dim b As Boolean

Sub StartCAB()
b = True
Sheets(1).Cells(8, "h") = Sheets(1).Cells(15, "j")
Do While b
Application.Wait (Now + #12:00:01 AM#)
DoEvents

If Sheets(1).Cells(8, "h") = #12:00:01 AM# Then MsgBox "Beep Beep Beep, le temps est écoulé !"
If Sheets(1).Cells(8, "h") = #12:00:01 AM# Then Exit Sub
Sheets(1).Cells(8, "h") = Format(DateAdd("s", -1, Sheets(1).Cells(8, "h")), "hh:mm:ss")

Loop
End Sub

Sub PauseCAB()
b = False
End Sub

Sub StopCAB()
b = False
Sheets(1).Cells(8, "h") = Sheets(1).Cells(15, "j")
End Sub

