Attribute VB_Name = "Module1"
Option Explicit

Sub BuscarYColocarEnHoja()
    Dim ws As Worksheet
    Dim buscarTexto As String
    Dim celdaEncontrada As Range
    Dim nuevaHoja As Worksheet
    Dim filaDestino As Integer
    
    ' Establecer el texto a buscar
    buscarTexto = "TOTAL TMU"
    
    ' Crear una nueva hoja para almacenar los resultados
    Set nuevaHoja = Sheets.Add(After:=Sheets(Sheets.Count))
    nuevaHoja.Name = "Resultados"
    
    ' Inicializar la fila de destino en la nueva hoja
    filaDestino = 1
    
    ' Recorrer todas las hojas del libro
    For Each ws In ThisWorkbook.Sheets
        ' Buscar el texto en la hoja actual
        Set celdaEncontrada = ws.Cells.Find(What:=buscarTexto, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Verificar si se encontró el texto
        If Not celdaEncontrada Is Nothing Then
            ' Obtener el valor de la celda a la derecha
            Dim valorDerecha As Variant
            valorDerecha = celdaEncontrada.Offset(0, 1).Value

           ' Colocar el nombre del archivo en la nueva hoja
            nuevaHoja.Cells(filaDestino, 1).Value = Thisworkbook.Name
            
            ' Colocar el nombre de la hoja en la nueva hoja
            nuevaHoja.Cells(filaDestino, 2).Value = ws.Name
            
            ' Colocar el valor a la derecha en la nueva hoja
            nuevaHoja.Cells(filaDestino, 3).Value = valorDerecha
            
            ' Incrementar la fila de destino en la nueva hoja
            filaDestino = filaDestino + 1
        End If
    Next ws
End Sub

