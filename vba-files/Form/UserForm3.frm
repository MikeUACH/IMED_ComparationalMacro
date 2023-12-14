VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Inicio"
   ClientHeight    =   9690.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17820
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim PathBU1 As String
Dim PathBU2 As String


Private Sub UserForm_Initialize()
    ActualizarEstadoBoton
End Sub ' Hacer un If que haga que seleccione archivo con nombre similar

Private Sub btnSeleccionarBU1_Click()
    If PathBU1 = "" Or PathBU1 = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathBU1 = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo BU Scenario Flexline (Primero a comparar)")
        
        ' Verifica si se seleccion� un archivo
        If PathBU = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
        
        ' Verifica si el nombre del archivo contiene la cadena "BU"
        If Not FileNameContainsStr(PathBU1, "BU") Then
            MsgBox "Por favor, selecciona un archivo cuyo nombre contenga 'BU'", vbExclamation
            PathBU = ""
            Exit Sub
        End If
    End If
    
    Debug.Print "Path BU1: " & PathBU1
    ActualizarEstadoBoton
End Sub

Private Sub btnSeleccionarBU2_Click()
    If PathBU2 = "" Or PathBU2 = "False" Then
        ' Abre el cuadro de di�logo de selecci�n de archivo
        PathBU2 = Application.GetOpenFilename("Archivos Excel (*.xlsb), *.xlsb", , "Selecciona el archivo BU Scenario Flexline (Segundo a comparar)")
        
        ' Verifica si se seleccion� un archivo
        If PathBU = "False" Then
            Exit Sub ' Si no se seleccion� un archivo, sale del procedimiento
        End If
        
        ' Verifica si el nombre del archivo contiene la cadena "BU"
        If Not FileNameContainsStr(PathBU2, "BU") Then
            MsgBox "Por favor, selecciona un archivo cuyo nombre contenga 'BU'", vbExclamation
            PathBU = ""
            Exit Sub
        End If
    End If
    
    Debug.Print "Path BU2: " & PathBU2
    ActualizarEstadoBoton
End Sub
    
Private Sub btnActualizar_Click()
    Dim wsUbicaciones As Worksheet
    Set wsUbicaciones = ThisWorkbook.Sheets("UbicacionesGuardadas")

    respuestaUbisEspecificas = MsgBox("�Quieres usar las ubicaciones guardadas en la hoja 'UbicacionesGuardadas'?", vbYesNo + vbExclamation, "Advertencia")
    If respuestaUbisEspecificas = vbYes Then
        ' Verificar si las ubicaciones no están vacías
        If (Len(wsUbicaciones.Range("B3").Value) > 0 Or Len(wsUbicaciones.Range("B4").Value) > 0) And usarUbicacionesEspecificas = False Then
            Dim respuestaUbicaciones As VbMsgBoxResult
            respuestaUbicaciones = MsgBox("Se usaran las ubicaciones proporcionadas por la hoja 'UbicacionesGuardadas'. �Deseas continuar?", vbYesNo + vbExclamation, "Advertencia")
            If respuestaUbicaciones = vbYes Then
                ' Hacer un If para verificar que se han seleccionado los archivos necesarios
                UpdWCstaffShiftTabsBU wsUbicaciones.Range("B4").Value, wsUbicaciones.Range("B3").Value  ' Llama al mï¿½dulo 1
                UpdWCellTabBU wsUbicaciones.Range("B5").Value, wsUbicaciones.Range("B3").Value   ' Llama al mï¿½dulo 3
                
                Dim wsRegistro As Worksheet
                Set wsRegistro = ThisWorkbook.Sheets("RegistroAcciones")
                Dim lastRow As Long
                lastRow = wsRegistro.Cells(wsRegistro.Rows.Count, "A").End(xlUp).Row + 1
                wsRegistro.Cells(lastRow, 1).Value = Now
                wsRegistro.Cells(lastRow, 2).Value = "Accion realizada en BU Scenario Flexline 1.x.xlsb"
                wsRegistro.Columns("A:B").AutoFit
            Else
                MsgBox "OperaciOn cancelada."
            End If
        End If
    Else
        respuestaUbisHojas = MsgBox("Se usaran las ubicaciones proporcionadas en los botones. �Deseas continuar?", vbYesNo + vbExclamation, "Advertencia")
        If respuestaUbisHojas = vbYes Then
            ' Verificar si las ubicaciones de los botones están vacías
            If Len(PathBU1) = 0 Or Len(PathBU2) = 0 Then
                MsgBox "Por favor, selecciona todas las ubicaciones en los botones antes de actualizar.", vbExclamation, "Advertencia"
                Exit Sub
            End If
            If Len(PathBU1) > 0 Or Len(PathBU2) > 0 Then
                ' El usuario ha hecho clic en Sí, proceder con la operación
                UpdWCstaffShiftTabsBU PathDL, PathBU  ' Llama al m?dulo 1
                UpdWCellTabBU PathWC, PathBU   ' Llama al m?dulo 3
            End If
        Else
            MsgBox "Operacion cancelada."
        End If
    End If
End Sub

Private Sub btnGenerarReporte_Click()
   ObtenerYColocarTabsUnabFlex PathFlex, PathVariance ' Llama al m�dulo 8
End Sub

Private Sub btnBorrarUbicacionBU1_Click()
    If Len(PathBU1) > 0 Then
        PathBU1 = "False"
        If PathBU1 = "False" Then
            Dim comprobarBU1 As Boolean
            comprobarBU1 = True
        End If

        If comprobarBU1 = True Then
            PathBU1 = ""
            MsgBox "Se ha borrado con exito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path BU1: " & PathBU1
    End If

    If Len(PathBU1) = 0 And comprobarBU1 = False Then
        MsgBox "No hay ningun archivo seleccionado"
    End If
End Sub

Private Sub btnBorrarUbicacionDL_Click()
    If Len(PathBU2) > 0 Then
        PathBU2 = "False"
        If PathBU2 = "False" Then
            Dim comprobarBU2 As Boolean
            comprobarBU2 = True
        End If

        If comprobarBU2 = True Then
            PathDL = ""
            MsgBox "Se ha borrado con exito"
            ActualizarEstadoBoton
        End If

        Debug.Print "Path BU2: " & PathBU2
    End If

    If Len(PathBU2) = 0 And comprobarBU2 = False Then
        MsgBox "No hay ningun archivo seleccionado"
    End If
End Sub

Private Sub btnBorrarUbicaciones_Click()
    Dim wsUbicaciones As Worksheet
    Set wsUbicaciones = ThisWorkbook.Sheets("UbicacionesGuardadas")
    wsUbicaciones.Range("B3:B4").Value = ""
    wsUbicaciones.Range("B3:B4").Interior.Color = RGB(255, 172, 172) ' Rojo
    ThisWorkbook.Sheets("UbicacionesGuardadas").Columns("B").AutoFit
End Sub

Private Sub btnGuardarUbicaciones_Click()
    Dim ubicacionesGuardadas As String
    Dim wsUbicaciones As Worksheet
    Set wsUbicaciones = ThisWorkbook.Sheets("UbicacionesGuardadas")

    If Len(PathBU1) > 0 Then
        ThisWorkbook.Sheets("UbicacionesGuardadas").Range("B3").Value = PathBU
        ubicacionesGuardadas = ubicacionesGuardadas & "BU1, "
    End If

    If Len(PathBU2) > 0 Then
        ThisWorkbook.Sheets("UbicacionesGuardadas").Range("B4").Value = PathDL
        ubicacionesGuardadas = ubicacionesGuardadas & "BU1, "
    End If

    If Len(wsUbicaciones.Range("B3").Value) > 0 Then
        wsUbicaciones.Range("B3").Interior.Color = RGB(171, 255, 174) ' Verde
    Else
        wsUbicaciones.Range("B3").Interior.Color = RGB(255, 172, 172) ' Rojo
    End If

    If Len(wsUbicaciones.Range("B4").Value) > 0 Then
        wsUbicaciones.Range("B4").Interior.Color = RGB(171, 255, 174) ' Verde
    Else
        wsUbicaciones.Range("B4").Interior.Color = RGB(255, 172, 172) ' Rojo
    End If

    If ubicacionesGuardadas = "False" Then
        ubicacionesGuardadas = ""
    End If
    
    If Len(ubicacionesGuardadas) > 0 Then
        MsgBox "Ubicaciones guardadas con exito: " & Left(ubicacionesGuardadas, Len(ubicacionesGuardadas) - 2)
    Else
        MsgBox "No hay ubicaciones para guardar"
    End If

    ThisWorkbook.Sheets("UbicacionesGuardadas").Columns("B").AutoFit
End Sub
    
Private Sub ActualizarEstadoBoton()
    If Len(PathBU1) = 0 Or PathBU1 = "False" Then
        btnSeleccionarBU1.BackColor = RGB(255, 172, 172)
        txtNotSelectedBU1.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarBU1.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoBU1 As String
        nombreArchivoBU1 = Mid(PathBU1, InStrRev(PathBU1, "\") + 1)
        txtNotSelectedBU1.Caption = "Seleccionado: " & nombreArchivoBU1
    End If
    
    If Len(PathBU2) = 0 Or PathBU2 = "False" Then
        btnSeleccionarBU2.BackColor = RGB(255, 172, 172)
        txtNotSelectedBU2.Caption = "No se ha seleccionado"
    Else
        btnSeleccionarBU2.BackColor = RGB(171, 255, 174)
        Dim nombreArchivoBU2 As String
        nombreArchivoBU2 = Mid(PathBU2, InStrRev(PathBU2, "\") + 1)
        txtNotSelectedBU2.Caption = "Seleccionado: " & nombreArchivoBU2
    End If
End Sub

Function FileNameContainsStr(filePath As String, strToFind As String) As Boolean
    ' Obtiene solo el nombre del archivo de la ruta completa
    Dim fileName As String
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))

    ' Comprueba si el nombre del archivo contiene la cadena proporcionada
    FileNameContainsStr = (InStr(1, fileName, strToFind, vbTextCompare) > 0)
End Function
