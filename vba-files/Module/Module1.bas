Attribute VB_Name = "Module1"
Dim archivoOrigenPath As String
Sub UpdWCstaffShiftTabsBU(ByVal archivoOrigenPath As String, ByVal ArchivoDestinoPath As String)
    Dim ArchivoDestino As Workbook
    Dim archivoOrigen As Workbook
    Dim hojaOrigen As Worksheet
    Dim turno As Integer
    
    ' Abre el archivo de origen seleccionado
    Set archivoOrigen = Workbooks.Open(archivoOrigenPath)
    ' Abre el archivo de destino seleccionado
    Set ArchivoDestino = Workbooks.Open(ArchivoDestinoPath)
    
    ' Define la hoja de c�lculo en el archivo de origen
    Set hojaOrigen = archivoOrigen.Sheets("IMED DL Breakdow")
    
    ' Definir los nombres de los turnos
    Dim nombresTurnos() As String
    nombresTurnos = Split("FirstShift,SecondShift,ThirdShift,FourTwentyShift,FourTwentyOneShift,FourTwentyTwoShift,FourTwentyThreeShift", ",")
    
    ' Procesar cada turno
    For turno = 1 To 7
        Dim turnoNombre As String
        turnoNombre = nombresTurnos(turno - 1)
        
        ' Buscar la coincidencia en hojaOrigen
        ArchivoDestino.Sheets("WCStaff Format").Range("S" & (45 + ((turno - 1) * 41)) & ":AD" & (81 + ((turno - 1) * 41))).Value = hojaOrigen.Range("S" & (45 + ((turno - 1) * 41)) & ":AD" & (81 + ((turno - 1) * 41))).Value
    Next turno
    
    ' Cerrar el archivo de origen sin guardar cambios
    archivoOrigen.Close SaveChanges:=False
End Sub



