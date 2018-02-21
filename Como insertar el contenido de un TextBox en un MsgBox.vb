 Como insertar el contenido de un TextBox en un MsgBox

Hola, que tal.

No sé si es esto excatamente lo que necesitas:


Private Sub Command1_Click()
If MsgBox("Por Alguna Razón NO esta COMPLETO el costo de la copia: " & Text1.Text & vbCrLf & "¿Quiere Completarlo? ", vbInformation + vbYesNo + vbDefaultButton1, "A T E N C I O N") = vbYes Then MsgBox "Hola" Else MsgBox "No Hola"
End Sub

Private Sub Form_Load()
Text1.Text = "Prueba de mensaje"
Range("A3").Interior.ColorIndex = 3
End Sub