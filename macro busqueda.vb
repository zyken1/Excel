Function insertaRegistro()

End Function



Private Function fileExisteRegistro(noIdentificacion As String,rangoConsulta As String) As Long
	Dim numeroFila As Long

	numeroFila = 0

	With Reto40Excel.Range(rangoConsulta)
		Set c = .Find(noIdentificacion, looKln = xlValues)

		if Not c Is Nothing Then
			numeroFila = c.Row
		End If
	End With

	'sale de la funcion
	fileExisteRegistro = numeroFila

End Function



With Worksheets(1).Range("a1:a500") 
    Set c = .Find(2, lookin:=xlValues) 
    If Not c Is Nothing Then 
        firstAddress = c.Address 
        Do 
            c.Value = 5 
            Set c = .FindNext(c) 
        Loop While Not c Is Nothing And c.Address <> firstAddress 
    End If 
End With