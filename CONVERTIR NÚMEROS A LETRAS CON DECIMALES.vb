 CONVERTIR NÚMEROS A LETRAS CON DECIMALES en VBA para Excel y hasta miles de billones.
Esta función convierte a letras tus números con decimales y hasta miles de billones.
La función es "num2let". 
Copia lo siguiente en VBA:

Function num2let(value)
  If Int(value) = 1 Then
  num2let = "uno" & " con " & Int(Round(((value - Int(value)) * 100))) & "/100"
  Else
   num2let = Num2Text(value) & " con " & Int(Round(((value - Int(value)) * 100))) & "/100"
End If
End Function


Public Function Num2Text(ByVal value As Double) As String
fraccion = value - Int(value)
value = Int(value)
    Select Case value
        Case 0: Num2Text = "cero"
        Case 1: Num2Text = "un"
        Case 2: Num2Text = "dos"
        Case 3: Num2Text = "tres"
        Case 4: Num2Text = "cuatro"
        Case 5: Num2Text = "cinco"
        Case 6: Num2Text = "seis"
        Case 7: Num2Text = "siete"
        Case 8: Num2Text = "ocho"
        Case 9: Num2Text = "nueve"
        Case 10: Num2Text = "diez"
        Case 11: Num2Text = "once"
        Case 12: Num2Text = "doce"
        Case 13: Num2Text = "trece"
        Case 14: Num2Text = "catorce"
        Case 15: Num2Text = "quince"
        Case Is < 20: Num2Text = "dieci" & Num2Text(value - 10)
        Case 20: Num2Text = "veinte"
        Case Is < 30: Num2Text = "veinti" & Num2Text(value - 20)
        Case 30: Num2Text = "treinta"
        Case 40: Num2Text = "cuarenta"
        Case 50: Num2Text = "cincuenta"
        Case 60: Num2Text = "sesenta"
        Case 70: Num2Text = "setenta"
        Case 80: Num2Text = "ochenta"
        Case 90: Num2Text = "noventa"
        Case Is < 100: Num2Text = Num2Text(Int(value \ 10) * 10) & " y " & Num2Text(value Mod 10)
        Case 100: Num2Text = "cien"
        Case Is < 200: Num2Text = "ciento " & Num2Text(value - 100)
        Case 200, 300, 400, 600, 800: Num2Text = Num2Text(Int(value \ 100)) & "cientos"
        Case 500: Num2Text = "quinientos"
        Case 700: Num2Text = "setecientos"
        Case 900: Num2Text = "novecientos"
        Case Is < 1000: Num2Text = Num2Text(Int(value \ 100) * 100) & " " & Num2Text(value Mod 100)
        Case 1000: Num2Text = "mil"
        Case Is < 2000: Num2Text = "mil " & Num2Text(value Mod 1000)
        Case Is < 1000000: Num2Text = Num2Text(Int(value \ 1000)) & " mil"
            If value Mod 1000 Then Num2Text = Num2Text & " " & Num2Text(value Mod 1000)
        Case 1000000: Num2Text = "un millón"
        Case Is < 2000000: Num2Text = "un millón " & Num2Text(value Mod 1000000)
        Case Is < 1000000000000#: Num2Text = Num2Text(Int(value / 1000000)) & " millones"
            If (value - Int(value / 1000000) * 1000000) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000) * 1000000)
        Case 1000000000000#: Num2Text = "un billón"
        Case Is < 2000000000000#: Num2Text = "un billón " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)
        Case Else: Num2Text = Num2Text(Int(value / 1000000000000#)) & " billones"
            If (value - Int(value / 1000000000000#) * 1000000000000#) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000000000#) * 1000000000000#)
    End Select
  End Function