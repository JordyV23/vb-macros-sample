Attribute VB_Name = "Module1"
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

    ' Ruta del archivo origen
    rutaArchivoOrigen = ""
    
    ' Columnas a eliminar (cambia las letras según tus necesidades)
    columnasAEliminar = Array("B", "D", "F") ' Ejemplo: eliminar columnas B, D, F
    
    ' Abre el archivo origen
    Set libroOrigen = Workbooks.Open(rutaArchivoOrigen)
    Set hojaOrigen = libroOrigen.Sheets(1) ' Cambia el número o nombre de la hoja
    
    ' Eliminar columnas innecesarias
    For i = UBound(columnasAEliminar) To LBound(columnasAEliminar) Step -1
        hojaOrigen.Columns(columnasAEliminar(i)).Delete
    Next i

    ' Eliminar duplicados (en función de la primera columna)
    ultimaColumna = hojaOrigen.Cells(1, hojaOrigen.Columns.Count).End(xlToLeft).Column
    ultimaFila = hojaOrigen.Cells(hojaOrigen.Rows.Count, 1).End(xlUp).Row + 1
    hojaOrigen.Range(hojaOrigen.Cells(2, 1), hojaOrigen.Cells(ultimaFila, ultimaColumna)).RemoveDuplicates Columns:=1, Header:=xlNo

    ' Copiar el contenido procesado (sin encabezados) al libro actual
    Set hojaDestino = ThisWorkbook.Sheets("") ' Cambia al número o nombre de la hoja destino en el libro actual
    ultimaFilaDestino = hojaDestino.Cells(hojaDestino.Rows.Count, 1).End(xlUp).Row + 1 ' Encuentra la última fila vacía en el destino
    hojaOrigen.Range("A2:G" & ultimaFila).Copy hojaDestino.Range("B" & ultimaFilaDestino)
    
    
    ' Copiar las fórmulas en columnas F y G a las nuevas filas
    If ultimaFilaDestino > 1 Then
        ' Asume que las fórmulas están en las primeras filas y las copia a las nuevas filas
        hojaDestino.Range("G2:H2").Copy hojaDestino.Range("G" & ultimaFilaDestino & ":H" & ultimaFilaDestino + ultimaFila - 1)
    End If
    
    ' Cierra el archivo origen sin guardar cambios
    libroOrigen.Close False

    MsgBox "El contenido ha sido procesado, pegado y las fórmulas se han aplicado correctamente."

End Sub
    
