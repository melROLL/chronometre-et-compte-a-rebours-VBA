Attribute VB_Name = "Chrono"
Dim a As Boolean

Sub StartChrono()
a = True 'ça commence le chrono'
  Do While a
     Application.Wait (Now + #12:00:01 AM#)
       DoEvents 'le code tourne pendant que l'on fait ça vie'
     Sheets(1).Cells(8, "C") = Format(DateAdd("s", 1, Sheets(1).Cells(8, "C")), "hh:mm:ss")
   Loop
End Sub

Sub PauseChrono()
    a = False 'ça stop le chrono'
End Sub

Sub StopChrono()
    a = False 'ça stop le chrono'
    Sheets(1).Cells(8, "C") = "00:00:00" 'on initialise le chrono'
End Sub

