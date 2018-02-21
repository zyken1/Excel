Sub CopyNoBlanks()
'este copia y da valores a la celda O y les da valor de numero
 ''
 Range("W5:W38").Select
    Selection.Copy
   ''
   Range("X5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   ''
    Range("X5").Select
    Application.CutCopyMode = False
    ''
    Range("X5:X38").Select
    Selection.TextToColumns Destination:=Range("X5"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'
'este es fotmato de copiar celdas escepto en blanco  aumenta una celda mas
'
Range("X5:X39").Copy Range("Y5")
''
With Range("Y5:Y38")
.Replace What:="0", Replacement:="", LookAt:=xlWhole
.Replace What:="", Replacement:="x"
.Replace What:="x", Replacement:=""
.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
End With
''
Range("Y5:Y39").Select
'Seañade un color negro ya que estaba pro defecto el gris
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
'Y se anexa por ultimo los valores a  HOJA DE COTIZACION
''
'Range("Y5:Y39").Select
    Selection.Copy
    Sheets("Hoja de Cotizacion").Select
    Range("A21").Select
    ActiveSheet.Paste
    Range("A21").Select
    Application.CutCopyMode = False
    
    
    Sheets("Cotizador ").Select
    Range("A5").Select
End Sub
