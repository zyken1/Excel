Private Sub Inversores_Click()
Sheets("Cotizador").Range("B7,C7,C8").Value = ""

      ' Dimension the variable.
      Dim x As Double
        x = Sheets("Cotizador").Range("F6").Value
     '  x = Sheets("RESUMEN").Range("C49").Value
      
      ' Display the value of x.
        MsgBox "La Potencia del Inversor es " & x & "."
      ' Start the Select Case structure.
      
    Select Case x
         
         Case Is <= 1.5
            'MsgBox "X is <=10"
             Sheets("Cotizador").Range("B6").Value = "GALVO"
             Sheets("Cotizador").Range("C6").Value = "Galvo- 1.5KW"
         Case Is <= 2
             Sheets("Cotizador").Range("B6").Value = "GALVO"
             Sheets("Cotizador").Range("C6").Value = "Galvo- 2.0KW"
         Case Is <= 2.5
             Sheets("Cotizador").Range("B6").Value = "GALVO"
             Sheets("Cotizador").Range("C6").Value = "Galvo-2.5 KW"
         Case Is <= 3.1
             Sheets("Cotizador").Range("B6").Value = "GALVO"
             Sheets("Cotizador").Range("C6").Value = "Galvo-3.1 KW"
         Case Is <= 3.8
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-3.8 KW"
         Case Is <= 5
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-5 KW"
         Case Is <= 6
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-6 KW"
         Case Is <= 7.6
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-7.6 KW"
         Case Is <= 8.2
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-8.2 KW"
         Case Is <= 10
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-10 KW"
             Sheets("Cotizador").Range("B7").Value = "SYSMO"
             Sheets("Cotizador").Range("C7").Value = "Symo-10.0/220"
         Case Is <= 11.4
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-11.4 KW"
         Case Is <= 12.5
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-12.5 KW"
             Sheets("Cotizador").Range("B7").Value = "SYSMO"
             Sheets("Cotizador").Range("C7").Value = "Symo-12.0/220"
         Case Is <= 15
             Sheets("Cotizador").Range("B6").Value = "PRIMO"
             Sheets("Cotizador").Range("C6").Value = "Primo-15 KW"
         Case Is <= 17.5
         
         Case Is <= 20
         
         Case Is <= 22.7
         
         Case Is <= 24
         
         Case Else
            MsgBox "X does not fall within the range"
    End Select
      
End Sub



'==============================================================================




Sub Inversores()

' Calculo para inversores Fronius Desde el calculo de modulos desde Celda  18 Hasta el 30
'

   'inicio:
   ' If Range("F5") < 1 Or Range("F5") >= 96 Then
   '  MsgBox ("La cantidad de modulos deben estar entre  1 y 96")
   ' End If
     'GoTo inicio

'Macro para boton Inversores
'Para obtener el calculo se multiplica Modulos * 260 y se obtiene el resultado
    If Range("F5").Value <= 6 Then
     Range("C7").Value = "1500"
    Else
       If Range("F5").Value <= 8 Then
        Range("C7").Value = "2000"
       Else
         If Range("F5").Value <= 10 Then
          Range("C7").Value = "2500"
         Else
           If Range("F5").Value <= 12 Then
             Range("C7").Value = "3100"
           Else
             If Range("F5").Value <= 14 Then
              Range("C7").Value = "3.8"
             Else
               If Range("F5").Value <= 20 Then
                Range("C7").Value = "5.0"
               Else
                 If Range("F5").Value <= 24 Then
                  Range("C7").Value = "6.0"
                 Else
                   If Range("F5").Value <= 28 Then
                    Range("C7").Value = "7.6"
                   Else
                     If Range("F5").Value <= 30 Then
                      Range("C7").Value = "8.2"
                     Else
                       If Range("F5").Value <= 36 Then
                        Range("C7").Value = "10.0"
                       Else
                         If Range("F5").Value <= 44 Then
                          Range("C7").Value = "11.4"
                         Else
                           If Range("F5").Value <= 48 Then
                            Range("C7").Value = "12.5"
                           Else
                             If Range("F5").Value <= 60 Then
                              Range("C7").Value = "15.0"
                             Else
                             End If
                           End If
                         End If
                       End If
                     End If
                    End If
                   End If
                 End If
               End If
            End If
         End If
      End If
    End If
    
'Inversores SYSMO
If Range("F5").Value <= 36 Then
    Range("C8").Value = "10.0/220"
   Else
      If Range("F5").Value <= 48 Then
       Range("C8").Value = "12.0/220"
      Else
         If Range("F5").Value <= 60 Then
          Range("C8").Value = "SYMO-15.0/220 (1MPPT)"
         Else
            If Range("F5").Value <= 50 Then
             Range("C8").Value = "SYMO-15.0/480"
            Else
             If Range("F5").Value <= 70 Then
              Range("C8").Value = "SYMO-17.5/480"
             Else
               If Range("F5").Value <= 80 Then
                Range("C8").Value = "SYMO-20.0/480"
               Else
                 If Range("F5").Value <= 90 Then
                  Range("C8").Value = "SYMO-22.7/480"
                 Else
                  If Range("F5").Value <= 96 Then
                   Range("C8").Value = "SYMO-24.0/480"
                  Else
                 End If
               End If
             End If
           End If
         End If
      End If
    End If
  End If

'Condicionales para el boton    PV Combiner
   If Range("F5").Value <= 1 Then
    Range("C9").Value = "PV Combiner 1"
   Else
      If Range("F5").Value <= 20 Then
       Range("C9").Value = "PV Combiner 2"
      Else
         If Range("F5").Value <= 30 Then
          Range("C9").Value = "PV Combiner 3"
         Else
         End If
      End If
   End If
    
End Sub




'==========================================================================0

=SI(RESUMEN!C49<=1.5,"Galvo-1.5KW",
SI(RESUMEN!C49<=2,"Galvo- 2.0KW",
SI(RESUMEN!C49<=2.5,"Galvo-2.5 KW",

Sheets("Cotizador").Range("C6").Value = "Galvo-3.1 KW"

SI(RESUMEN!C49<=3.8,"Primo-3.8 KW",
SI(RESUMEN!C49<=5,"Primo-5.0 KW",
SI(RESUMEN!C49<=6,"Primo-6.0 KW",
SI(RESUMEN!C49<=7.6,"Primo-7.6 KW",
SI(RESUMEN!C49<=8.2,"Primo-.8.2 KW ",S
I(RESUMEN!C49<=10,"Pirmo-10.0 KW",SI(

 Galvo-1.5KW 
 Galvo- 2.0KW 
 Galvo-2.5 KW 
 Galvo-3.1 KW 
 Primo-3.8 KW 
 Primo-5.0 KW 
 Primo-6.0 KW 
 Primo-7.6 KW 
 Primo-.8.2 KW 
 Pirmo-10.0 KW 
 Primo-11.4 KW 
 Primo-12.5 KW 
 Primo-15.0 KW 
 
 Symo-10.0/220 
 Symo-12.0/220 
 Symo-15.0/220 (1MPPT) 
 Symo-15.0/220 (1MPPT) 
 Symo-10.0/480 
 Symo-12.5/480 
 Symo-15.0/480 
 Symo-17.5/480 
 Symo-20.0/480 
 Symo-22.7/480 
 Symo-24.0/480 