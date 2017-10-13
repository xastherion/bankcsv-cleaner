Attribute VB_Name = "Module1"
Sub banco_csv_limpiar()

'borra las columnas segun letra
'------------------------------------------------------------------------
    Dim operacion As String
    Dim columna(9)
    Dim Val(9)
    
    Val(1) = "B"
    Val(2) = "C"
    Val(3) = "E"
    Val(4) = "F"
    Val(5) = "G"
    Val(6) = "H"
    Val(7) = "I"
    Val(8) = "J"
    Val(9) = "K"
    
    For contador = 9 To 1 Step -1
        operacion = Val(contador)
        
        'MsgBox (contador)
        'MsgBox (operacion)
        
        Range(operacion + "1").Select
        Columns(operacion).Select
        Selection.Delete
        Range("A1").Select
    
    Next
    
'borra primeras lineas
'-------------------------------------------------------------------------
    Range("A1:O12").Select
    Selection.EntireRow.Delete

'aplicar formato :
'-------------------------------------------------------------------------

    'selecciona columna A (Fecha) y le asigna formato fecha centrado
    Columns("A:A").Select
    Selection.NumberFormat = "m/d/yyyy"
    Selection.HorizontalAlignment = xlCenter
    
    'tamano columna B
    Columns("B:B").ColumnWidth = 66.57
   
    'columna de el monto, le asigna el formato moneda
    Columns("C:C").NumberFormat = _
        "_-* #,##0.00 [$€-de-DE]_-;-* #,##0.00 [$€-de-DE]_-;_-* ""-""?? [$€-de-DE]_-;_-@_-"
    
    'fila uno, en negrita
    Rows("1:1").Font.Bold = True

'columna Haben
'------------------------------

    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC[-1]=""H"",RC[-2],0)"
    Selection.AutoFill Destination:=Range("E2:E60"), Type:=xlFillDefault
    Range("E2:E60").Select
    ActiveWindow.SmallScroll Down:=-75

'columna SOLL
'--------------------------------
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=+IF(RC[-2]=""S"",RC[-3],0)"
    Selection.AutoFill Destination:=Range("F2:F60"), Type:=xlFillDefault
    Range("F2:F60").Select
    ActiveWindow.SmallScroll Down:=-42
   
'limpieza2
'-------------------------------
    Columns("C:D").Select
    Range("D1").Activate
    Selection.EntireColumn.Hidden = True
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Haben"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Soll"
    Columns("E:F").Select
    Selection.Style = "Currency"
    'congela linea superior
      With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    'tableformat
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$F$70"), , xlYes).Name = _
        "Table11"
    Range("Table11[#All]").Select
    ActiveSheet.ListObjects("Table11").TableStyle = "TableStyleLight1"
    'formato condicional para Haben y para Soll
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    
    Columns("F:F").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249946592608417
    'buchungs y emfanger formato condicional
    
    End With
    
End Sub



