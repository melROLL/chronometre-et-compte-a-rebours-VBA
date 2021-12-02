# Chronomètre, compte à rebours Excel VBA
## Les fonctionnalités
Cete Excel possède deux fonctionnalités, une première qui est un chronomètre et la seconde qui est un compte à rebours.
 
-	Le chronomètre est présent à gauche et possède 3 boutons ayant une fonctionnalité différente :
-	Le bouton play (à gauche) : ce bouton commence le chronomètre, celui-ci affiche le temps sous forme hh:mm:ss. Celui-ci s’incrémente même lorsque vous travaillez sur d’autres feuilles de calculs.
-	Le bouton pause (au milieu) : ce bouton arrête le chronomètre et affiche le temps qui a été compté par celui-ci. Lorsque vous appuyez à nouveau sur le bouton play, le temps s’ajoute à celui déjà écoulé.
-	Le bouton stop (à droite) : ce bouton arrête le chronomètre et le remet à 0 .

-	Le compte à rebours est présent à droite et possède 3 boutons possédant chacun une fonctionnalité différente :
-	Le bouton play (à gauche) : ce bouton commence le compte à rebours.  Le temps est affiché dans la cellule en dessous du bouton au format hh:mm:ss. Le décompte se fait à partir d’une valeur donnée.
-	Le bouton pause (au milieu) : ce bouton arrête le compte à rebours et laisse afficher le temps restant du compte à rebours. 	Lorsque vous appuyez à nouveau sur le bouton play, le décompte reprend là où il s’était arrêté.
-	Le bouton stop (à droite) : ce bouton arrête le décompte et remet le compte à rebours à la valeur donnée par l’utilisateur.
La valeur utilisée pour le compte à rebours est celle rentrée par l’utilisateur dans la cellule J15.

## Les sources utilisées
 - https://www.youtube.com/watch?v=XIIQMzE0BO8
 - https://fr.extendoffice.com/documents/excel/3684-excel-create-stopwatch.html

## Code VBA utilisé
Dans ce projet, nous avons utilisé 2 codes différents, composés de 3 parties chacun. Un premier code pour le chronomètre et un second pour le compte à rebours.

1.	Code Chronomètre :
```VBA
Dim a As Boolean
Sub StartChrono()
a = True 'ça commence le chrono'
  Do While a
     Application.Wait (Now + #12:00:01 AM#)
       DoEvents 'le code tourne pendant que l'on fait sa vie'
     Sheets(1).Cells(8, "C") = Format(DateAdd("s", 1, Sheets(1).Cells(8, "C")), "hh:mm:ss")
   Loop
End Sub
Sub PauseChrono()
    a = False 'ça stoppe le chrono'
End Sub
Sub StopChrono()
    a = False 'ça stoppe le chrono'
    Sheets(1).Cells(8, "C") = "00:00:00" 'on initialise le chrono'
End Sub
```

2.	Code compte à rebours :
```VBA
Dim b As Boolean
Sub StartCAB()
b = True ‘ça commence le compte à rebours’
Sheets(1).Cells(8, "h") = Sheets(1).Cells(15, "j")
Do While b
Application.Wait (Now + #12:00:01 AM#)
DoEvents
If Sheets(1).Cells(8, "h") = #12:00:01 AM# Then MsgBox "Beep Beep Beep, le temps est écoulé !"
If Sheets(1).Cells(8, "h") = #12:00:01 AM# Then Exit Sub
Sheets(1).Cells(8, "h") = Format(DateAdd("s", -1, Sheets(1).Cells(8, "h")), "hh:mm:ss")Loop
End Sub
Sub PauseCAB()
b = False ‘ça arrête le compte à rebours’
End Sub
Sub StopCAB()
b = False ‘ça arrête le compte à rebours’
Sheets(1).Cells(8, "h") = Sheets(1).Cells(15, "j")
End Sub
```
