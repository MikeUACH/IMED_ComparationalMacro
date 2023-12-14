Attribute VB_Name = "Module2"
Dim archivoOrigenPath As String

Sub UpdNonMatMarginBU(ByVal archivoOrigenPath As String, ByVal ArchivoDestinoPath As String)
    Dim ArchivoDestino As Workbook
    Dim TotalFlexline As Variant
    Dim hojaOrigen As Worksheet

    ' Abre el archivo de origen seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    
    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("AllocationTotal")
    
    ' Buscar la coincidencia en hojaOrigen
    TotalFlexline = hojaOrigen.Range("D59:O69")
    
    ' Coloca los valores obtenidos en celdas espec�ficas de tu hoja de c�lculo principal
    ArchivoDestino.Sheets("Non Mat Margin").Range("D168:O178").Value = TotalFlexline
End Sub





