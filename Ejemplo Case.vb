Macro de ejemplo que usa la instrucción Select Case

   Sub Using_Case()
      ' Dimension the variable.
      Dim x As Integer
      ' Place a value in x.
      x = Int(Rnd * 100)
      ' Display the value of x.
      MsgBox "The value of x is " & x & "."
      ' Start the Select Case structure.
      Select Case x
         ' Test to see if x less than or equal to 10.
         Case Is <= 10
            ' Display a message box.
            MsgBox "X is <=10"
         ' Test to see if x less than or equal to 40 and greater than 10.
         Case 11 To 40
            MsgBox "X is <=40 and > 10"
         ' Test to see if x less than or equal to 70 and greater than 40.
         Case 41 To 70
            MsgBox "X is <=70 and > 40"
         ' Test to see if x less than or equal to 100 and greater than 70.
         Case 71 To 100
            MsgBox "X is <= 100 and > 70"
         ' If none of the above tests returned true.
         Case Else
            MsgBox "X does not fall within the range"
      End Select
   End Sub
   
   
   '======================================================================
   
   
   Sub EjemploSelectCase()
Dim numero As Integer
numero = 8
Select Case numero
Case 1 To 5
     MsgBox “El número esta entre 1 y 5 “
Case 6, 7, 8
     MsgBox “El número esta entre 6, 7 y 8 “
Case 9 To 10
     MsgBox “El número esta entre 9 y 10 “
Case Else
     MsgBox “El número no esta entre 1 y 10 “
End Select

End Sub