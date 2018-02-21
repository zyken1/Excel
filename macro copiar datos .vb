
Private Sub botonGuardar_Click()
'===========BOTON GUARDAR
'===========ANTES SE LLAMABA DATA

ult = Sheets("Base de Expedientes").Cells(Rows.Count, 4).End(xlUp).Row

If noExpediente = "" Or partes = "" Or asunto = "" Then
    MsgBox "Escriba todos los datos"
Else
    
    Sheets("Base de Expedientes").Cells(ult + 1, 2) = noExpediente
    Sheets("Base de Expedientes").Cells(ult + 1, 3) = boxTipoExpediente
    Sheets("Base de Expedientes").Cells(ult + 1, 4) = partes
    Sheets("Base de Expedientes").Cells(ult + 1, 5) = asunto

    diseñoCeldas1 = Módulo1.diseñoCeldas    ' diseño de las celdas

   
noExpediente = ""
tipoExpediente = ""
partes = ""
asunto = ""


End If


End Sub