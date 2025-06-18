Attribute VB_Name = "Módulo1"
Sub Filtro_Suma_4X1000()

'Declaración de variables
Dim ws As Worksheet
Dim lastRow As Long
Dim celda As Range
Dim suma As Double
Dim valorLimpio As String

'Seleccionar la hoja datos
Set ws = ThisWorkbook.Sheets("datos")

'Quitar filtros previos
If ws.AutoFilterMode Then ws.AutoFilterMode = False

'Aplicar autofiltro en la fila 1 (cabecera)
lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
ws.Range("A1:F" & lastRow).AutoFilter Field:=2, Criteria1:="IMPTO GOBIERNO 4X1000"

'Inicializar la suma
suma = 0

'Recorrer las celdas visibles en la coumna D (columna 4)
Dim r As Range
On Error Resume Next
Set r = ws.Range("D2:D" & lastRow).SpecialCells(xlCellTypeVisible)
On Error GoTo 0

If Not r Is Nothing Then
    For Each celda In r
        If Not IsEmpty(celda.Value) Then
        
            'Eliminar puntos y comas
            valorLimpio = Replace(celda.Value, ".", "")
            valorLimpio = Replace(valorLimpio, ",", ".")
            
            'Verificar si es numerico despues de la limpieza
            If IsNumeric(valorLimpio) Then
                suma = suma + CDbl(valorLimpio)
            End If
        End If
    Next celda
End If

'Mostrar el resultado en H2 y aplicar formato moneda
ws.Range("H2").Value = suma
ws.Range("H1").NumberFormat = "$#,##0.00"  ' Formato de moneda con símbolo de dólar


' Quitar el filtro
ws.AutoFilterMode = False

MsgBox "Suma completada. Resultado en H2: " & Format(suma, "$#,##0.00")
            



End Sub
