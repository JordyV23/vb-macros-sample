Sub AbrirProcesarYCopiarConFormulas()

    Dim libroOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim hojaDestino As Worksheet
    Dim rutaArchivoOrigen As String
    Dim columnasAEliminar As Variant
    Dim ultimaColumna As Long
    Dim ultimaFila As Long
    Dim ultimaFilaDestino As Long
    Dim i As Long

    ' Ruta del archivo origen 1
    rutaArchivoOrigen = ""
    
    ' Ruta del archivo origen 2
    rutaArchivoOrigen2 = ""
    
    ' Ruta del archivo origen 3
    rutaArchivoOrigen3 = ""
    
    ' Ruta del archivo origen 4
    rutaArchivoOrigen4 = ""
    
    ' Columnas a eliminar (cambia las letras seg�n tus necesidades)
    columnasAEliminar = Array("B", "F") ' Ejemplo: eliminar columnas B, F
    
    ' Abre el archivo origen
    Set libroOrigen = Workbooks.Open(rutaArchivoOrigen)
    Set hojaOrigen = libroOrigen.Sheets(1) ' Cambia el n�mero o nombre de la hoja
    
    ' Eliminar columnas innecesarias
    For i = UBound(columnasAEliminar) To LBound(columnasAEliminar) Step -1
        hojaOrigen.Columns(columnasAEliminar(i)).Delete
    Next i

    ' Eliminar duplicados (en funci�n de la primera columna)
    ultimaColumna = hojaOrigen.Cells(2, hojaOrigen.Columns.Count).End(xlToLeft).Column
    ultimaFila = hojaOrigen.Cells(hojaOrigen.Rows.Count, 1).End(xlUp).Row + 1
    hojaOrigen.Range(hojaOrigen.Cells(2, 1), hojaOrigen.Cells(ultimaFila, ultimaColumna)).RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' Recalcular indices
    ultimaColumna = hojaOrigen.Cells(2, hojaOrigen.Columns.Count).End(xlToLeft).Column
    ultimaFila = hojaOrigen.Cells(hojaOrigen.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Copiar el contenido procesado (sin encabezados) al libro actual (solo columnas A-E)
    Set hojaDestino = ThisWorkbook.Sheets("Actual") ' Cambia al n�mero o nombre de la hoja destino en el libro actual
    ultimaFilaDestino = hojaDestino.Cells(hojaDestino.Rows.Count, 1).End(xlUp).Row + 1 ' Encuentra la �ltima fila vac�a en el destino
    hojaOrigen.Range("A2:G" & ultimaFila).Copy hojaDestino.Range("C" & ultimaFilaDestino)
    
    
    ' Copiar las f�rmulas en columnas F y G a las nuevas filas
    If ultimaFilaDestino > 1 Then
        ' Asume que las f�rmulas est�n en las primeras filas y las copia a las nuevas filas
        hojaDestino.Range("G2:H2").Copy hojaDestino.Range("G" & ultimaFilaDestino & ":H" & ultimaFilaDestino + ultimaFila - 1)
        
        ' Asume que las f�rmulas est�n en las primeras filas y las copia a las nuevas filas
        hojaDestino.Range("A2:B2").Copy hojaDestino.Range("A" & ultimaFilaDestino & ":B" & ultimaFilaDestino + ultimaFila - 1)
    End If
    
    ' Cierra el archivo origen sin guardar cambios
    libroOrigen.Close False
    

    MsgBox "El contenido ha sido procesado, pegado y las f�rmulas se han aplicado correctamente."

End Sub
    
