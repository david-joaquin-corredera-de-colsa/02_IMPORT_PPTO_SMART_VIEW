Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_01"
Option Explicit


Public Function fun801_LogMessage(ByVal strMessage As String, _
                                Optional ByVal blnIsError As Boolean = False, _
                                Optional ByVal strFileName As String = "", _
                                Optional ByVal strSheetName As String = "") As Boolean
        
    '------------------------------------------------------------------------------
    ' FUNCI�N: fun801_LogMessage
    ' PROP�SITO: Sistema integral de logging para registrar eventos y errores
    '
    ' PAR�METROS:
    ' - strMessage (String): Mensaje a registrar
    ' - blnIsError (Boolean, Opcional): True=ERROR, False=INFO (defecto: False)
    ' - strFileName (String, Opcional): Archivo relacionado (defecto: "NA")
    ' - strSheetName (String, Opcional): Hoja relacionada (defecto: "NA")
    '
    ' RETORNA: Boolean - True si exitoso, False si error
    '
    ' FUNCIONALIDADES:
    ' - Crea hoja de log autom�ticamente con formato profesional
    ' - Timestamp ISO, usuario del sistema, tipo de evento
    ' - Formato condicional para errores (fondo rojo)
    ' - Filtros autom�ticos y ajuste de columnas
    '
    ' COMPATIBILIDAD: Excel 97-365, Office Online, SharePoint, Teams
    '
    ' EJEMPLO: Call fun801_LogMessage("Operaci�n completada", False, "datos.csv")
    '------------------------------------------------------------------------------
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para el log
    Dim wsLog As Worksheet
    Dim lngLastRow As Long
    Dim strDateTime As String
    Dim strUserName As String
    Dim strLogType As String
    
    ' Inicializaci�n
    strFuncion = "fun801_LogMessage" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun801_LogMessage"
    fun801_LogMessage = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Verificar hoja de log (constant CONST_HOJA_LOG = "02_Log")
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If Not fun802_SheetExists(CONST_HOJA_LOG) Then
        If Not F002_Crear_Hoja(CONST_HOJA_LOG) Then
            MsgBox "Error al crear la hoja de log", vbCritical
            Exit Function
        End If
        
        ' Crear y formatear encabezados
        With ThisWorkbook.Sheets(CONST_HOJA_LOG)
            ' Establecer textos de encabezados exactamente como se solicita
            .Range("A1").Value = "Date/Time"
            .Range("B1").Value = "User"
            .Range("C1").Value = "Type"
            .Range("D1").Value = "File"
            .Range("E1").Value = "Sheet"
            .Range("F1").Value = "Message"
            
            ' Formato de encabezados
            With .Range("A1:F1")
                .Font.Bold = True
                .Font.Size = 11
                .Font.Name = "Calibri"
                .Interior.Color = RGB(200, 200, 200)
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlMedium
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            ' Formato espec�fico para la columna de fecha
            .Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
            
            ' Ajustar anchos de columna
            .Columns("A").ColumnWidth = 20  ' Date/Time
            .Columns("B").ColumnWidth = 15  ' User
            .Columns("C").ColumnWidth = 15  ' Type
            .Columns("D").ColumnWidth = 40  ' File
            .Columns("E").ColumnWidth = 20  ' Sheet
            .Columns("F").ColumnWidth = 60  ' Message
            
            ' Filtros autom�ticos
            .Range("A1:F1").AutoFilter
        End With
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Preparar datos para el log
    '--------------------------------------------------------------------------
    lngLineaError = 55
    Set wsLog = ThisWorkbook.Sheets(CONST_HOJA_LOG)
    
    ' Obtener �ltima fila
    lngLastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Preparar datos (reemplazar valores vac�os con "NA")
    strDateTime = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    strUserName = IIf(Environ("USERNAME") = "", "NA", Environ("USERNAME"))
    strLogType = IIf(blnIsError, "ERROR", "INFO")
    strFileName = IIf(Len(Trim(strFileName)) = 0, "NA", strFileName)
    strSheetName = IIf(Len(Trim(strSheetName)) = 0, "NA", strSheetName)
    strMessage = IIf(Len(Trim(strMessage)) = 0, "NA", strMessage)
    
    '--------------------------------------------------------------------------
    ' 3. Escribir en el log
    '--------------------------------------------------------------------------
    lngLineaError = 70
    With wsLog
        ' Escribir datos
        .Cells(lngLastRow, 1).Value = strDateTime    ' Date/Time
        .Cells(lngLastRow, 2).Value = strUserName    ' User
        .Cells(lngLastRow, 3).Value = strLogType     ' Type
        .Cells(lngLastRow, 4).Value = strFileName    ' File
        .Cells(lngLastRow, 5).Value = strSheetName   ' Sheet
        .Cells(lngLastRow, 6).Value = strMessage     ' Message
        
        ' Formato de la nueva fila
        With .Range(.Cells(lngLastRow, 1), .Cells(lngLastRow, 6))
            ' Formato general
            .Font.Name = "Calibri"
            .Font.Size = 10
            .VerticalAlignment = xlTop
            .WrapText = True
            
            ' Bordes
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            
            ' Formato condicional para errores
            If blnIsError Then
                .Interior.Color = RGB(255, 200, 200)
                .Font.Bold = True
            End If
        End With
        
        ' Asegurar formato de fecha en la columna A
        .Cells(lngLastRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    
    fun801_LogMessage = True
    Exit Function

GestorErrores:
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    MsgBox strMensajeError, vbCritical, "Error en sistema de logging"
    fun801_LogMessage = False
End Function




Public Function F002_Crear_Hoja(ByVal strNombreHoja As String) As Boolean

    '******************************************************************************
    ' M�dulo: F002_Crear_Hoja
    ' Fecha y Hora de Creaci�n: 2025-05-26 09:17:15 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Funci�n para crear hojas en el libro con formato y configuraci�n est�ndar
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para manejo de hojas
    Dim ws As Worksheet
    
    ' Inicializaci�n
    strFuncion = "F002_Crear_Hoja" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F002_Crear_Hoja"
    F002_Crear_Hoja = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Verificar si la hoja ya existe
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If fun802_SheetExists(strNombreHoja) Then
        F002_Crear_Hoja = True
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Crear nueva hoja
    '--------------------------------------------------------------------------
    lngLineaError = 40
    Application.ScreenUpdating = False
    
    ' Crear hoja al final del libro
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    
    ' Asignar nombre
    ws.Name = strNombreHoja
    
    ' Configuraci�n b�sica
    'With ws
    '    ' Ajustar vista
    '    .DisplayGridlines = True
    '    .DisplayHeadings = True
    '
    '    ' Configurar primera vista
    '    .Range("A1").Select
    '
    '    ' Ajustar ancho de columnas est�ndar
    '    .Columns.StandardWidth = 10
    '
    '    ' Configurar �rea de impresi�n
    '    .PageSetup.PrintArea = ""
    'End With
    
    Application.ScreenUpdating = True
    
    F002_Crear_Hoja = True
    Exit Function

GestorErrores:
    ' Restaurar configuraci�n
    Application.ScreenUpdating = True
    
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F002_Crear_Hoja = False
End Function



Public Function fun801_LimpiarHoja(ByVal strNombreHoja As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun801_LimpiarHoja
    ' FECHA Y HORA DE CREACI�N: 2025-05-28 17:50:26 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa    '
    ' PROP�SITO:
    ' Limpia de forma segura y eficiente todo el contenido de una hoja de c�lculo
    ' espec�fica, preservando el formato y estructura, pero eliminando todos los
    ' datos y valores almacenados en las celdas utilizadas.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(strNombreHoja)
    
    Application.ScreenUpdating = False
    ws.UsedRange.ClearContents
    Application.ScreenUpdating = True
    
    fun801_LimpiarHoja = True
    Exit Function
    
GestorErrores:
    fun801_LimpiarHoja = False
End Function

Public Function fun802_SeleccionarArchivo(ByVal strPrompt As String) As String
    
    '******************************************************************************
    ' FUNCI�N: fun802_SeleccionarArchivo (VERSI�N MEJORADA)
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' FECHA MODIFICACI�N: 2025-06-01
    '
    ' PROP�SITO:
    ' Proporciona una interfaz de usuario intuitiva para seleccionar archivos de
    ' texto (TXT y CSV) con sistema de carpetas de respaldo autom�tico.
    '
    ' L�GICA DE CARPETAS DE RESPALDO:
    ' 1. Carpeta del archivo Excel actual
    ' 2. %TEMP% (si hay error)
    ' 3. %TMP% (si hay error)
    ' 4. %USERPROFILE% (si hay error)
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para carpetas de respaldo
    Dim strCarpetaInicial As String
    Dim strCarpetaActual As String
    Dim intIntentoActual As Integer
    Dim blnCarpetaValida As Boolean
    
    ' Inicializaci�n
    strFuncion = "fun802_SeleccionarArchivo" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun802_SeleccionarArchivo"
    fun802_SeleccionarArchivo = ""
    lngLineaError = 0
    intIntentoActual = 1
    blnCarpetaValida = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Intentar obtener carpetas de respaldo en orden de prioridad
    '--------------------------------------------------------------------------
    Do While intIntentoActual <= 4 And Not blnCarpetaValida
        lngLineaError = 40 + intIntentoActual
        
        Select Case intIntentoActual
            Case 1: ' Carpeta del archivo Excel actual
                strCarpetaActual = fun803_ObtenerCarpetaExcelActual()
                
            Case 2: ' Variable de entorno %TEMP%
                strCarpetaActual = fun804_ObtenerCarpetaTemp()
                
            Case 3: ' Variable de entorno %TMP%
                strCarpetaActual = fun805_ObtenerCarpetaTmp()
                
            Case 4: ' Variable de entorno %USERPROFILE%
                strCarpetaActual = fun806_ObtenerCarpetaUserProfile()
        End Select
        
        ' Verificar si la carpeta es v�lida y accesible
        If fun807_ValidarCarpeta(strCarpetaActual) Then
            blnCarpetaValida = True
            strCarpetaInicial = strCarpetaActual
        Else
            intIntentoActual = intIntentoActual + 1
        End If
    Loop
    
    ' Si no se pudo obtener ninguna carpeta v�lida, usar carpeta por defecto
    If Not blnCarpetaValida Then
        strCarpetaInicial = ""
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Mostrar di�logo de selecci�n de archivo
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    On Error GoTo GestorErrores
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = strPrompt
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt;*.csv"
        .AllowMultiSelect = False
        
        ' Establecer carpeta inicial si es v�lida
        If Len(strCarpetaInicial) > 0 Then
            .InitialFileName = strCarpetaInicial & "\"
        End If
        
        If .Show = -1 Then
            fun802_SeleccionarArchivo = .SelectedItems(1)
        Else
            fun802_SeleccionarArchivo = ""
        End If
    End With
    
    Exit Function
    
GestorErrores:
    ' Log del error
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Intento actual: " & intIntentoActual
    
    fun801_LogMessage strMensajeError, True
    fun802_SeleccionarArchivo = ""
End Function

Public Function fun803_ImportarArchivo(ByRef wsDestino As Worksheet, _
                                     ByVal strFilePath As String, _
                                     ByVal strColumnaInicial As String, _
                                     ByVal lngFilaInicial As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N: fun803_ImportarArchivo
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Importa el contenido completo de archivos de texto plano (TXT/CSV) l�nea por
    ' l�nea hacia una hoja de Excel espec�fica, colocando cada l�nea del archivo
    ' en una celda individual seg�n la posici�n inicial definida. Funci�n core
    ' del sistema de importaci�n de datos de presupuesto.
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim objFSO As Object
    Dim objFile As Object
    Dim strLine As String
    Dim lngRow As Long
    
    ' Inicializaci�n
    strFuncion = "fun803_ImportarArchivo" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun803_ImportarArchivo"
    fun803_ImportarArchivo = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros
    '--------------------------------------------------------------------------
    lngLineaError = 20
    If wsDestino Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, "Hoja de destino no v�lida"
    End If
    
    If Len(strFilePath) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, "Ruta de archivo no v�lida"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Configurar objetos para lectura de archivo
    '--------------------------------------------------------------------------
    lngLineaError = 35
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFilePath, 1) ' ForReading = 1
    
    '--------------------------------------------------------------------------
    ' 3. Leer archivo l�nea por l�nea
    '--------------------------------------------------------------------------
    lngLineaError = 45
    lngRow = lngFilaInicial
    
    While Not objFile.AtEndOfStream
        strLine = objFile.ReadLine
        wsDestino.Range(strColumnaInicial & lngRow).Value = strLine
        lngRow = lngRow + 1
    Wend
    
    '--------------------------------------------------------------------------
    ' 4. Limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 60
    objFile.Close
    Set objFile = Nothing
    Set objFSO = Nothing
    
    fun803_ImportarArchivo = True
    Exit Function

GestorErrores:
    ' Limpieza en caso de error
    If Not objFile Is Nothing Then
        objFile.Close
        Set objFile = Nothing
    End If
    Set objFSO = Nothing
    
    ' Construcci�n del mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True
    fun803_ImportarArchivo = False
End Function


Public Function fun804_DetectarRangoDatos(ByRef ws As Worksheet, _
                                         ByRef lngLineaInicial As Long, _
                                         ByRef lngLineaFinal As Long) As Boolean
    '******************************************************************************
    ' FUNCI�N: fun804_DetectarRangoDatos
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Detecta autom�ticamente el rango exacto de datos en una columna espec�fica
    ' de una hoja de c�lculo, identificando la primera y �ltima fila que contienen
    ' informaci�n. Funci�n esencial para determinar l�mites de procesamiento
    ' despu�s de la importaci�n de datos.
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim rngBusqueda As Range
    Dim lngColumna As Long
    
    ' Obtener n�mero de columna
    lngColumna = Range(vColumnaInicial_Importacion & "1").Column
    
    ' Configurar rango de b�squeda
    Set rngBusqueda = ws.Columns(lngColumna)
    
    With rngBusqueda
        ' Encontrar primera celda con datos
        Set rngBusqueda = .Find(What:="*", _
                               After:=.Cells(.Cells.Count), _
                               LookIn:=xlFormulas, _
                               LookAt:=xlPart, _
                               SearchOrder:=xlByRows, _
                               SearchDirection:=xlNext)
        
        If Not rngBusqueda Is Nothing Then
            lngLineaInicial = rngBusqueda.Row
            
            ' Encontrar �ltima celda con datos
            Set rngBusqueda = .Find(What:="*", _
                                   After:=.Cells(1), _
                                   LookIn:=xlFormulas, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByRows, _
                                   SearchDirection:=xlPrevious)
            
            lngLineaFinal = rngBusqueda.Row
            fun804_DetectarRangoDatos = True
        Else
            lngLineaInicial = 0
            lngLineaFinal = 0
            fun804_DetectarRangoDatos = False
        End If
    End With
    Exit Function
    
GestorErrores:
    lngLineaInicial = 0
    lngLineaFinal = 0
    fun804_DetectarRangoDatos = False
End Function




Public Function fun801_VerificarExistenciaHoja(wb As Workbook, nombreHoja As String) As Boolean
    ' =============================================================================
    ' FUNCI�N AUXILIAR 801: VERIFICAR EXISTENCIA DE HOJA
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Verifica si una hoja existe en el libro especificado
    ' Par�metros: wb (Workbook), nombreHoja (String)
    ' Retorna: Boolean (True si existe, False si no existe)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim lineaError As Long
    
    lineaError = 200
    fun801_VerificarExistenciaHoja = False
    
    ' Verificar par�metros de entrada
    If wb Is Nothing Or nombreHoja = "" Then
        Exit Function
    End If
    
    lineaError = 210
    
    ' Recorrer todas las hojas del libro (m�todo compatible con Excel 97)
    For i = 1 To wb.Worksheets.Count
        If UCase(wb.Worksheets(i).Name) = UCase(nombreHoja) Then
            fun801_VerificarExistenciaHoja = True
            Exit For
        End If
    Next i
    
    lineaError = 220
    
    Exit Function
    
ErrorHandler:
    fun801_VerificarExistenciaHoja = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun801_VerificarExistenciaHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "PAR�METRO nombreHoja: " & nombreHoja & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function



Public Sub fun804_LimpiarContenidoHoja(ws As Worksheet)
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 804: LIMPIAR CONTENIDO DE HOJA
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Limpia todo el contenido de una hoja espec�fica
    ' Par�metros: ws (Worksheet)
    ' Retorna: Nada (Sub procedure)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 500
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        Exit Sub
    End If
    
    lineaError = 510
    
    ' Verificar que la hoja no est� protegida
    If ws.ProtectContents Then
        ws.Unprotect
    End If
    
    lineaError = 520
    
    ' Limpiar todo el contenido de la hoja (m�todo compatible con todas las versiones)
    ws.Cells.Clear
    
    lineaError = 530
    
    Exit Sub
    
ErrorHandler:
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun804_LimpiarContenidoHoja" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Sub


Public Function fun805_DetectarUseSystemSeparators() As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 805: DETECTAR USE SYSTEM SEPARATORS
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta si Excel est� usando separadores del sistema
    ' Par�metros: Ninguno
    ' Retorna: String ("True" o "False")
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variable para almacenar el resultado
    Dim resultado As String
    Dim lineaError As Long
    
    lineaError = 600
    
    ' Detectar configuraci�n actual de Use System Separators
    ' Usar compilaci�n condicional para compatibilidad con versiones
    
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lineaError = 610
        If Application.UseSystemSeparators Then
            resultado = "True"
        Else
            resultado = "False"
        End If
    #Else
        ' Excel 97, 2003 y anteriores
        lineaError = 620
        resultado = fun809_DetectarUseSystemSeparatorsLegacy()
    #End If
    
    lineaError = 630
    
    fun805_DetectarUseSystemSeparators = resultado
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, intentar m�todo alternativo
    fun805_DetectarUseSystemSeparators = fun809_DetectarUseSystemSeparatorsLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun805_DetectarUseSystemSeparators" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun806_DetectarDecimalSeparator() As String

    ' =============================================================================
    ' FUNCI�N AUXILIAR 806: DETECTAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta el separador decimal actual de Excel
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador decimal)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 700
    
    ' Detectar separador decimal actual (compatible con todas las versiones)
    fun806_DetectarDecimalSeparator = Application.DecimalSeparator
    
    lineaError = 710
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar m�todo alternativo
    fun806_DetectarDecimalSeparator = fun810_DetectarDecimalSeparatorLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun806_DetectarDecimalSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun807_DetectarThousandsSeparator() As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 807: DETECTAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: Detecta el separador de miles actual de Excel
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador de miles)
    ' Compatibilidad: Excel 97, 2003, 365
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 800
    
    ' Detectar separador de miles actual (compatible con todas las versiones)
    fun807_DetectarThousandsSeparator = Application.ThousandsSeparator
    
    lineaError = 810
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, usar m�todo alternativo
    fun807_DetectarThousandsSeparator = fun811_DetectarThousandsSeparatorLegacy()
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun807_DetectarThousandsSeparator" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function





Public Function fun809_DetectarUseSystemSeparatorsLegacy() As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 809: DETECTAR USE SYSTEM SEPARATORS (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar Use System Separators en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String ("True" o "False")
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para comparaci�n
    Dim separadorSistema As String
    Dim separadorExcel As String
    Dim lineaError As Long
    
    lineaError = 1000
    
    ' Obtener separador decimal del sistema (Windows)
    ' M�todo compatible con Excel 97 y 2003
    separadorSistema = Mid(CStr(1.1), 2, 1)
    
    lineaError = 1010
    
    ' Obtener separador decimal de Excel
    separadorExcel = Application.DecimalSeparator
    
    lineaError = 1020
    
    ' Si coinciden, probablemente Use System Separators est� activado
    If separadorSistema = separadorExcel Then
        fun809_DetectarUseSystemSeparatorsLegacy = "True"
    Else
        fun809_DetectarUseSystemSeparatorsLegacy = "False"
    End If
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir False por defecto
    fun809_DetectarUseSystemSeparatorsLegacy = "False"
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun809_DetectarUseSystemSeparatorsLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun810_DetectarDecimalSeparatorLegacy() As String
    ' =============================================================================
    ' FUNCI�N AUXILIAR 810: DETECTAR DECIMAL SEPARATOR (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar separador decimal en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador decimal)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detecci�n
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1100
    
    ' M�todo alternativo: formatear un n�mero y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = CStr(1.1)
    
    lineaError = 1110
    
    ' El separador decimal es el segundo car�cter en el formato est�ndar
    If Len(numeroFormateado) >= 2 Then
        fun810_DetectarDecimalSeparatorLegacy = Mid(numeroFormateado, 2, 1)
    Else
        ' Fallback: asumir punto por defecto
        fun810_DetectarDecimalSeparatorLegacy = "."
    End If
    
    lineaError = 1120
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir punto por defecto
    fun810_DetectarDecimalSeparatorLegacy = "."
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun810_DetectarDecimalSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function

Public Function Inventario_Actualizado_Si_No() As Boolean
    
    '******************************************************************************
    ' FUNCI�N: Inventario_Actualizado_Si_No
    ' FECHA Y HORA DE CREACI�N: 2025-01-15 14:30:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Compara el estado actual de las hojas del libro con la informaci�n almacenada
    ' en la hoja de inventario para determinar si el inventario est� actualizado.
    ' Verifica tanto la existencia de hojas como su estado de visibilidad.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables y configuraci�n de optimizaci�n
    ' 2. Recopilaci�n de informaci�n actual de todas las hojas del libro
    ' 3. Lectura de informaci�n del inventario desde la hoja correspondiente
    ' 4. Comparaci�n bidireccional entre arrays de hojas existentes e inventariadas
    ' 5. Validaci�n de concordancia en nombres y estados de visibilidad
    ' 6. Generaci�n de logging detallado de discrepancias encontradas
    ' 7. Restauraci�n de configuraci�n y retorno del resultado
    '
    ' PAR�METROS: Ninguno
    ' RETORNA: Boolean - True si inventario actualizado, False si hay discrepancias
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimizaci�n
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas y datos
    Dim ws As Worksheet
    Dim wsInventario As Worksheet
    Dim lngTotalHojasLibro As Long
    Dim lngContadorHojas As Long
    Dim lngUltimaFilaInventario As Long
    Dim lngFilaActual As Long
    
    ' Arrays para almacenar informaci�n
    Dim vHojasExistentes() As Variant
    Dim vHojasInventariadas() As Variant
    Dim vNumeroHojasInventariadas As Integer
    Dim lngContadorInventario As Long
    
    ' Variables para comparaci�n y validaci�n
    Dim strNombreHoja As String
    Dim blnHojaVisible As Boolean
    Dim strValorColumnaVisible As String
    Dim blnEncontrado As Boolean
    Dim lngIndiceComparacion As Long
    
    ' Inicializaci�n
    strFuncion = "Inventario_Actualizado_Si_No" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Inventario_Actualizado_Si_No"
    Inventario_Actualizado_Si_No = False
    lngLineaError = 0
    vNumeroHojasInventariadas = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables y configuraci�n de optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    Call fun801_LogMessage("Iniciando verificaci�n de actualizaci�n del inventario", False, "", strFuncion)
    
    ' Guardar configuraci�n original para restaurar despu�s
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    ' Configurar optimizaci�n de rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 2. Recopilaci�n de informaci�n actual de todas las hojas del libro
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Obtener n�mero total de hojas en el libro
    lngTotalHojasLibro = ThisWorkbook.Worksheets.Count
    
    If lngTotalHojasLibro = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9001, strFuncion, _
            "No hay hojas en el libro de trabajo"
    End If
    
    ' Dimensionar array para hojas existentes (2 dimensiones)
    ReDim vHojasExistentes(1 To lngTotalHojasLibro, 1 To 2)
    
    Call fun801_LogMessage("Recopilando informaci�n de " & lngTotalHojasLibro & " hojas existentes", _
        False, "", strFuncion)
    
    ' Recorrer todas las hojas del libro y recopilar informaci�n
    For lngContadorHojas = 1 To lngTotalHojasLibro
        lngLineaError = 70 + lngContadorHojas
        
        Set ws = ThisWorkbook.Worksheets(lngContadorHojas)
        
        ' Almacenar nombre de la hoja (dimensi�n 1)
        vHojasExistentes(lngContadorHojas, 1) = ws.Name
        
        ' Almacenar estado de visibilidad (dimensi�n 2)
        ' True si visible, False si oculta
        vHojasExistentes(lngContadorHojas, 2) = (ws.Visible = xlSheetVisible)
        
        Call fun801_LogMessage("Hoja " & lngContadorHojas & ": " & Chr(34) & ws.Name & Chr(34) & _
            " - Visible: " & CStr(vHojasExistentes(lngContadorHojas, 2)), False, "", strFuncion)
    Next lngContadorHojas
    
    '--------------------------------------------------------------------------
    ' 3. Lectura de informaci�n del inventario desde la hoja correspondiente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Verificar existencia de hoja de inventario
    If Not fun802_SheetExists(CONST_HOJA_INVENTARIO) Then
        Err.Raise ERROR_BASE_IMPORT + 9002, strFuncion, _
            "La hoja de inventario no existe: " & CONST_HOJA_INVENTARIO
    End If
    
    Set wsInventario = ThisWorkbook.Worksheets(CONST_HOJA_INVENTARIO)
    
    ' Encontrar �ltima fila con datos en la columna de nombres
    lngUltimaFilaInventario = wsInventario.Cells(wsInventario.Rows.Count, CONST_INVENTARIO_COLUMNA_NOMBRE).End(xlUp).Row
    
    Call fun801_LogMessage("�ltima fila con datos en inventario: " & lngUltimaFilaInventario, _
        False, "", strFuncion)
    
    ' Verificar que hay datos despu�s de los headers
    If lngUltimaFilaInventario <= CONST_INVENTARIO_FILA_HEADERS Then
        Call fun801_LogMessage("WARNING: No hay datos en el inventario despu�s de la fila de headers", _
            True, "", strFuncion)
        GoTo RestaurarConfiguracion ' Considerar como no actualizado
    End If
    
    '--------------------------------------------------------------------------
    ' 3.1. Contar hojas inventariadas (con datos v�lidos)
    '--------------------------------------------------------------------------
    lngLineaError = 110
    
    vNumeroHojasInventariadas = 0
    
    ' Recorrer filas del inventario para contar las que tienen nombre de hoja
    For lngFilaActual = CONST_INVENTARIO_FILA_HEADERS + 1 To lngUltimaFilaInventario
        strNombreHoja = Trim(CStr(wsInventario.Cells(lngFilaActual, CONST_INVENTARIO_COLUMNA_NOMBRE).Value))
        
        If Len(strNombreHoja) > 0 Then
            vNumeroHojasInventariadas = vNumeroHojasInventariadas + 1
        End If
    Next lngFilaActual
    
    Call fun801_LogMessage("N�mero de hojas inventariadas con datos v�lidos: " & vNumeroHojasInventariadas, _
        False, "", strFuncion)
    
    If vNumeroHojasInventariadas = 0 Then
        Call fun801_LogMessage("WARNING: No hay hojas inventariadas con datos v�lidos", _
            True, "", strFuncion)
        GoTo RestaurarConfiguracion ' Considerar como no actualizado
    End If
    
    '--------------------------------------------------------------------------
    ' 3.2. Llenar array de hojas inventariadas
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
    ' Dimensionar array para hojas inventariadas
    ReDim vHojasInventariadas(1 To vNumeroHojasInventariadas, 1 To 2)
    
    lngContadorInventario = 0
    
    ' Llenar array con datos del inventario
    For lngFilaActual = CONST_INVENTARIO_FILA_HEADERS + 1 To lngUltimaFilaInventario
        lngLineaError = 130 + lngFilaActual
        
        strNombreHoja = Trim(CStr(wsInventario.Cells(lngFilaActual, CONST_INVENTARIO_COLUMNA_NOMBRE).Value))
        
        If Len(strNombreHoja) > 0 Then
            lngContadorInventario = lngContadorInventario + 1
            
            ' Almacenar nombre de hoja (dimensi�n 1)
            vHojasInventariadas(lngContadorInventario, 1) = strNombreHoja
            
            ' Obtener y transformar valor de visibilidad (dimensi�n 2)
            strValorColumnaVisible = Trim(CStr(wsInventario.Cells(lngFilaActual, CONST_INVENTARIO_COLUMNA_VISIBLE).Value))
            
            ' Transformar seg�n especificaciones:
            ' "OCULTA" -> False (hoja oculta)
            ' ">> visible <<" -> True (hoja visible)
            If StrComp(strValorColumnaVisible, "OCULTA", vbTextCompare) = 0 Then
                vHojasInventariadas(lngContadorInventario, 2) = False
            ElseIf StrComp(strValorColumnaVisible, ">> visible <<", vbTextCompare) = 0 Then
                vHojasInventariadas(lngContadorInventario, 2) = True
            Else
                ' Valor no reconocido, asumir visible por defecto y registrar warning
                vHojasInventariadas(lngContadorInventario, 2) = True
                Call fun801_LogMessage("WARNING: Valor de visibilidad no reconocido para hoja " & Chr(34) & _
                    strNombreHoja & Chr(34) & ": " & Chr(34) & strValorColumnaVisible & Chr(34) & _
                    ". Asumiendo visible.", True, "", strFuncion)
            End If
            
            Call fun801_LogMessage("Inventario " & lngContadorInventario & ": " & Chr(34) & strNombreHoja & _
                Chr(34) & " - Visible: " & CStr(vHojasInventariadas(lngContadorInventario, 2)), _
                False, "", strFuncion)
        End If
    Next lngFilaActual
    
    '--------------------------------------------------------------------------
    ' 4. Comparaci�n bidireccional entre arrays
    '--------------------------------------------------------------------------
    lngLineaError = 200
    
    Call fun801_LogMessage("Iniciando comparaci�n bidireccional de arrays", False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 4.1. Verificar que cada hoja existente est� en el inventario
    '--------------------------------------------------------------------------
    lngLineaError = 210
    
    For lngContadorHojas = 1 To lngTotalHojasLibro
        lngLineaError = 220 + lngContadorHojas
        
        strNombreHoja = CStr(vHojasExistentes(lngContadorHojas, 1))
        blnHojaVisible = CBool(vHojasExistentes(lngContadorHojas, 2))
        blnEncontrado = False
        
        ' Buscar la hoja actual en el inventario
        For lngIndiceComparacion = 1 To vNumeroHojasInventariadas
            If StrComp(CStr(vHojasInventariadas(lngIndiceComparacion, 1)), strNombreHoja, vbTextCompare) = 0 Then
                blnEncontrado = True
                
                ' Comparar estado de visibilidad
                If CBool(vHojasInventariadas(lngIndiceComparacion, 2)) <> blnHojaVisible Then
                    Call fun801_LogMessage("DISCREPANCIA: Hoja " & Chr(34) & strNombreHoja & Chr(34) & _
                        " - Estado actual: " & CStr(blnHojaVisible) & _
                        ", Estado en inventario: " & CStr(vHojasInventariadas(lngIndiceComparacion, 2)), _
                        True, "", strFuncion)
                    GoTo RestaurarConfiguracion ' Retornar False
                End If
                Exit For
            End If
        Next lngIndiceComparacion
        
        ' Si la hoja no se encontr� en el inventario
        If Not blnEncontrado Then
            Call fun801_LogMessage("DISCREPANCIA: Hoja existente " & Chr(34) & strNombreHoja & _
                Chr(34) & " no encontrada en el inventario", True, "", strFuncion)
            GoTo RestaurarConfiguracion ' Retornar False
        End If
    Next lngContadorHojas
    
    '--------------------------------------------------------------------------
    ' 4.2. Verificar que cada hoja inventariada existe realmente
    '--------------------------------------------------------------------------
    lngLineaError = 250
    
    For lngContadorInventario = 1 To vNumeroHojasInventariadas
        lngLineaError = 260 + lngContadorInventario
        
        strNombreHoja = CStr(vHojasInventariadas(lngContadorInventario, 1))
        blnHojaVisible = CBool(vHojasInventariadas(lngContadorInventario, 2))
        blnEncontrado = False
        
        ' Buscar la hoja inventariada en las hojas existentes
        For lngIndiceComparacion = 1 To lngTotalHojasLibro
            If StrComp(CStr(vHojasExistentes(lngIndiceComparacion, 1)), strNombreHoja, vbTextCompare) = 0 Then
                blnEncontrado = True
                
                ' Comparar estado de visibilidad
                If CBool(vHojasExistentes(lngIndiceComparacion, 2)) <> blnHojaVisible Then
                    Call fun801_LogMessage("DISCREPANCIA: Hoja inventariada " & Chr(34) & strNombreHoja & _
                        Chr(34) & " - Estado en inventario: " & CStr(blnHojaVisible) & _
                        ", Estado actual: " & CStr(vHojasExistentes(lngIndiceComparacion, 2)), _
                        True, "", strFuncion)
                    GoTo RestaurarConfiguracion ' Retornar False
                End If
                Exit For
            End If
        Next lngIndiceComparacion
        
        ' Si la hoja inventariada no existe realmente
        If Not blnEncontrado Then
            Call fun801_LogMessage("DISCREPANCIA: Hoja inventariada " & Chr(34) & strNombreHoja & _
                Chr(34) & " no existe en el libro actual", True, "", strFuncion)
            GoTo RestaurarConfiguracion ' Retornar False
        End If
    Next lngContadorInventario
    
    '--------------------------------------------------------------------------
    ' 5. Si llegamos aqu�, el inventario est� actualizado
    '--------------------------------------------------------------------------
    lngLineaError = 300
    
    Call fun801_LogMessage("�XITO: El inventario est� completamente actualizado. " & _
        "Hojas existentes: " & lngTotalHojasLibro & ", Hojas inventariadas: " & vNumeroHojasInventariadas, _
        False, "", strFuncion)
    
    Inventario_Actualizado_Si_No = True

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 6. Restauraci�n de configuraci�n y limpieza
    '--------------------------------------------------------------------------
    lngLineaError = 350
    
    ' Restaurar configuraci�n original
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set ws = Nothing
    Set wsInventario = Nothing
    
    Call fun801_LogMessage("Verificaci�n de inventario completada. Resultado: " & _
        CStr(Inventario_Actualizado_Si_No), False, "", strFuncion)
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 7. Manejo exhaustivo de errores
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hojas en libro: " & lngTotalHojasLibro & vbCrLf & _
                      "Hojas inventariadas: " & vNumeroHojasInventariadas & vbCrLf & _
                      "Hoja actual procesando: " & strNombreHoja & vbCrLf & _
                      "Fecha y Hora: " & Now()
    
    ' Registrar error en log
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Mostrar error al usuario (opcional)
    MsgBox strMensajeError, vbCritical, "Error en Verificaci�n de Inventario"
    
    ' Restaurar configuraci�n en caso de error
    On Error Resume Next
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias
    Set ws = Nothing
    Set wsInventario = Nothing
    
    ' Retornar False en caso de error
    Inventario_Actualizado_Si_No = False
End Function
' =============================================================================
' FUNCION: Ordenar_Hojas
' FECHA: 2025-06-13 08:28:44 UTC
' DESCRIPCION: Ordena las pesta�as del libro con prioridad por visibilidad y formato de nombre
' PARAMETROS: Ninguno
' RETORNO: Boolean (True=�xito, False=error)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function Ordenar_Hojas() As Boolean

    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Optimizar configuraci�n de Excel para mejor rendimiento
    ' 2. Recopilar informaci�n de todas las hojas del libro
    ' 3. Separar hojas visibles y ocultas en arrays independientes
    ' 4. Categorizar cada grupo por patr�n de nombre (con/sin prefijo num�rico)
    ' 5. Ordenar lexicogr�ficamente cada subcategor�a por separado
    ' 6. Reorganizar las hojas seg�n el orden establecido
    ' 7. Restaurar configuraci�n original de Excel
    ' 8. Retornar resultado de la operaci�n

    On Error GoTo ErrorHandler
    
    Dim vResultado As Boolean
    Dim vLineaError As Integer
    Dim vTotalHojas As Integer
    Dim vContadorHojas As Integer
    Dim vNombreHoja As String
    Dim vEsVisible As Boolean
    
    ' Arrays para almacenar hojas visibles categorizadas
    Dim vHojasVisiblesConPrefijo() As String
    Dim vHojasVisiblesSinPrefijo() As String
    Dim vNumVisiblesConPrefijo As Integer
    Dim vNumVisiblesSinPrefijo As Integer
    
    ' Arrays para almacenar hojas ocultas categorizadas
    Dim vHojasOcultasConPrefijo() As String
    Dim vHojasOcultasSinPrefijo() As String
    Dim vNumOcultasConPrefijo As Integer
    Dim vNumOcultasSinPrefijo As Integer
    
    ' Variables para ordenamiento y control
    Dim i As Integer, j As Integer
    Dim vTempNombre As String
    Dim vPosicionActual As Integer
    
    ' Variables para optimizaci�n (inicializaci�n correcta)
    Dim vCalculationOriginal As Integer
    Dim vScreenUpdatingOriginal As Boolean
    Dim vEnableEventsOriginal As Boolean
    
    ' Variables para manejo de alertas
    Dim vDisplayAlertsOriginal As Boolean
    
    ' Inicializaci�n de variables
    vResultado = False
    vLineaError = 10
    vNumVisiblesConPrefijo = 0
    vNumVisiblesSinPrefijo = 0
    vNumOcultasConPrefijo = 0
    vNumOcultasSinPrefijo = 0
    vPosicionActual = 1
    
    ' Paso 1: Optimizar configuraci�n de Excel para mejor rendimiento
    vLineaError = 20
    
    ' Guardar configuraci�n original
    vCalculationOriginal = Application.Calculation
    vScreenUpdatingOriginal = Application.ScreenUpdating
    vEnableEventsOriginal = Application.EnableEvents
    vDisplayAlertsOriginal = Application.DisplayAlerts
    
    ' Aplicar optimizaciones
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Registrar inicio de operaci�n en log (con control de errores)
    On Error Resume Next
    Call fun801_LogMessage("Iniciando ordenamiento avanzado de hojas", False, "", "Ordenar_Hojas")
    On Error GoTo ErrorHandler
    
    ' Paso 2: Recopilar informaci�n de todas las hojas del libro
    vLineaError = 30
    vTotalHojas = ThisWorkbook.Worksheets.Count
    
    ' Validar que hay hojas para procesar
    If vTotalHojas <= 1 Then
        vResultado = True ' No hay nada que ordenar, pero no es error
        GoTo RestaurarConfiguracion
    End If
    
    ' Redimensionar arrays con tama�o m�ximo posible
    ReDim vHojasVisiblesConPrefijo(1 To vTotalHojas)
    ReDim vHojasVisiblesSinPrefijo(1 To vTotalHojas)
    ReDim vHojasOcultasConPrefijo(1 To vTotalHojas)
    ReDim vHojasOcultasSinPrefijo(1 To vTotalHojas)
    
    ' Paso 3: Separar hojas visibles y ocultas en arrays independientes
    ' Paso 4: Categorizar cada grupo por patr�n de nombre
    vLineaError = 40
    For vContadorHojas = 1 To vTotalHojas
        vNombreHoja = ThisWorkbook.Worksheets(vContadorHojas).Name
        vEsVisible = (ThisWorkbook.Worksheets(vNombreHoja).Visible = xlSheetVisible)
        
        If vEsVisible Then
            ' Hoja visible: categorizar por patr�n de nombre
            If fun801_TienePrefijoNumerico(vNombreHoja) Then
                vNumVisiblesConPrefijo = vNumVisiblesConPrefijo + 1
                vHojasVisiblesConPrefijo(vNumVisiblesConPrefijo) = vNombreHoja
            Else
                vNumVisiblesSinPrefijo = vNumVisiblesSinPrefijo + 1
                vHojasVisiblesSinPrefijo(vNumVisiblesSinPrefijo) = vNombreHoja
            End If
        Else
            ' Hoja oculta: categorizar por patr�n de nombre
            If fun801_TienePrefijoNumerico(vNombreHoja) Then
                vNumOcultasConPrefijo = vNumOcultasConPrefijo + 1
                vHojasOcultasConPrefijo(vNumOcultasConPrefijo) = vNombreHoja
            Else
                vNumOcultasSinPrefijo = vNumOcultasSinPrefijo + 1
                vHojasOcultasSinPrefijo(vNumOcultasSinPrefijo) = vNombreHoja
            End If
        End If
    Next vContadorHojas
    
    ' Paso 5: Ordenar lexicogr�ficamente cada subcategor�a por separado
    vLineaError = 50
    
    ' Ordenar hojas visibles con prefijo num�rico
    If vNumVisiblesConPrefijo > 1 Then
        For i = 1 To vNumVisiblesConPrefijo - 1
            For j = 1 To vNumVisiblesConPrefijo - i
                If StrComp(vHojasVisiblesConPrefijo(j), vHojasVisiblesConPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasVisiblesConPrefijo(j)
                    vHojasVisiblesConPrefijo(j) = vHojasVisiblesConPrefijo(j + 1)
                    vHojasVisiblesConPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Ordenar hojas visibles sin prefijo num�rico
    If vNumVisiblesSinPrefijo > 1 Then
        For i = 1 To vNumVisiblesSinPrefijo - 1
            For j = 1 To vNumVisiblesSinPrefijo - i
                If StrComp(vHojasVisiblesSinPrefijo(j), vHojasVisiblesSinPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasVisiblesSinPrefijo(j)
                    vHojasVisiblesSinPrefijo(j) = vHojasVisiblesSinPrefijo(j + 1)
                    vHojasVisiblesSinPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Ordenar hojas ocultas con prefijo num�rico
    If vNumOcultasConPrefijo > 1 Then
        For i = 1 To vNumOcultasConPrefijo - 1
            For j = 1 To vNumOcultasConPrefijo - i
                If StrComp(vHojasOcultasConPrefijo(j), vHojasOcultasConPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasOcultasConPrefijo(j)
                    vHojasOcultasConPrefijo(j) = vHojasOcultasConPrefijo(j + 1)
                    vHojasOcultasConPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Ordenar hojas ocultas sin prefijo num�rico
    If vNumOcultasSinPrefijo > 1 Then
        For i = 1 To vNumOcultasSinPrefijo - 1
            For j = 1 To vNumOcultasSinPrefijo - i
                If StrComp(vHojasOcultasSinPrefijo(j), vHojasOcultasSinPrefijo(j + 1), vbTextCompare) > 0 Then
                    vTempNombre = vHojasOcultasSinPrefijo(j)
                    vHojasOcultasSinPrefijo(j) = vHojasOcultasSinPrefijo(j + 1)
                    vHojasOcultasSinPrefijo(j + 1) = vTempNombre
                End If
            Next j
        Next i
    End If
    
    ' Paso 6: Reorganizar las hojas seg�n el orden establecido
    vLineaError = 60
    
    ' 6.1: Primero las hojas visibles con prefijo num�rico
    For i = 1 To vNumVisiblesConPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasVisiblesConPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    ' 6.2: Despu�s las hojas visibles sin prefijo num�rico
    For i = 1 To vNumVisiblesSinPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasVisiblesSinPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    ' 6.3: Despu�s las hojas ocultas con prefijo num�rico
    For i = 1 To vNumOcultasConPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasOcultasConPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    ' 6.4: Finalmente las hojas ocultas sin prefijo num�rico
    For i = 1 To vNumOcultasSinPrefijo
        Call fun803_Mover_Hoja_A_Posicion_Segura(vHojasOcultasSinPrefijo(i), vPosicionActual)
        vPosicionActual = vPosicionActual + 1
    Next i
    
    vResultado = True
    
RestaurarConfiguracion:
    ' Paso 7: Restaurar configuraci�n original de Excel
    vLineaError = 70
    On Error Resume Next
    Application.DisplayAlerts = vDisplayAlertsOriginal
    Application.EnableEvents = vEnableEventsOriginal
    Application.ScreenUpdating = vScreenUpdatingOriginal
    Application.Calculation = vCalculationOriginal
    On Error GoTo 0
    
    ' Registrar finalizaci�n en log (con control de errores)
    If vResultado Then
        On Error Resume Next
        Call fun801_LogMessage("Ordenamiento de hojas completado exitosamente. Total procesadas: " & _
            CStr(vTotalHojas) & ", Visibles con prefijo: " & CStr(vNumVisiblesConPrefijo) & _
            ", Visibles sin prefijo: " & CStr(vNumVisiblesSinPrefijo) & _
            ", Ocultas con prefijo: " & CStr(vNumOcultasConPrefijo) & _
            ", Ocultas sin prefijo: " & CStr(vNumOcultasSinPrefijo), False, "", "Ordenar_Hojas")
        On Error GoTo 0
    End If
    
    ' Paso 8: Retornar resultado de la operaci�n
    Ordenar_Hojas = vResultado
    Exit Function
    
ErrorHandler:
    Dim vMensajeError As String
    vMensajeError = "ERROR en Ordenar_Hojas" & vbCrLf & _
                   "Linea aproximada: " & vLineaError & vbCrLf & _
                   "Numero de Error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & _
                   "Usuario: david-joaquin-corredera-de-colsa" & vbCrLf & _
                   "Fecha y Hora: 2025-06-13 08:28:44 UTC"
    
    ' Restaurar configuraci�n en caso de error
    On Error Resume Next
    Application.DisplayAlerts = vDisplayAlertsOriginal
    Application.EnableEvents = vEnableEventsOriginal
    Application.ScreenUpdating = vScreenUpdatingOriginal
    Application.Calculation = vCalculationOriginal
    On Error GoTo 0
    
    ' Registrar error en log
    On Error Resume Next
    Call fun801_LogMessage(vMensajeError, True, "", "Ordenar_Hojas")
    On Error GoTo 0
    
    MsgBox vMensajeError, vbCritical, "Error Ordenar_Hojas"
    
    Ordenar_Hojas = False
    
End Function

' =============================================================================
' FUNCION AUXILIAR: fun801_TienePrefijoNumerico
' FECHA: 2025-06-13 08:28:44 UTC
' DESCRIPCION: Verifica si el nombre de hoja tiene prefijo con formato "##_"
' PARAMETROS: vNombreHoja (String)
' RETORNO: Boolean (True si tiene prefijo num�rico, False si no)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Function fun801_TienePrefijoNumerico(vNombreHoja As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim vPrimerCaracter As String
    Dim vSegundoCaracter As String
    Dim vTercerCaracter As String
    
    ' Inicializaci�n
    fun801_TienePrefijoNumerico = False
    
    ' Verificar que el nombre tenga al menos 3 caracteres
    If Len(vNombreHoja) < 3 Then
        Exit Function
    End If
    
    ' Extraer los primeros tres caracteres
    vPrimerCaracter = Mid(vNombreHoja, 1, 1)
    vSegundoCaracter = Mid(vNombreHoja, 2, 1)
    vTercerCaracter = Mid(vNombreHoja, 3, 1)
    
    ' Verificar patr�n: dos d�gitos seguidos de gui�n bajo
    ' Usar verificaci�n manual para compatibilidad con Excel 97
    If (vPrimerCaracter >= "0" And vPrimerCaracter <= "9") And _
       (vSegundoCaracter >= "0" And vSegundoCaracter <= "9") And _
       vTercerCaracter = Chr(95) Then
        fun801_TienePrefijoNumerico = True
    End If
    
    Exit Function
    
ErrorHandler:
    fun801_TienePrefijoNumerico = False
    
End Function

' =============================================================================
' SUB AUXILIAR: fun803_Mover_Hoja_A_Posicion_Segura
' FECHA: 2025-06-13 08:28:44 UTC
' DESCRIPCION: Mueve una hoja a una posici�n espec�fica con control de errores
' PARAMETROS: vNombreHoja (String), vPosicion (Integer)
' COMPATIBILIDAD: Excel 97-365, OneDrive/SharePoint/Teams
' =============================================================================
Public Sub fun803_Mover_Hoja_A_Posicion_Segura(vNombreHoja As String, vPosicion As Integer)
    
    On Error GoTo ErrorHandler
    
    Dim vHoja As Worksheet
    Dim vTotalHojas As Integer
    Dim vPosicionActualHoja As Integer
    Dim vHojaReferencia As Worksheet
    
    ' Verificar que la posici�n es v�lida
    vTotalHojas = ThisWorkbook.Worksheets.Count
    If vPosicion < 1 Or vPosicion > vTotalHojas Then
        Exit Sub
    End If
    
    ' Verificar que la hoja existe
    Set vHoja = Nothing
    On Error Resume Next
    Set vHoja = ThisWorkbook.Worksheets(vNombreHoja)
    On Error GoTo ErrorHandler
    
    If vHoja Is Nothing Then
        Exit Sub
    End If
    
    vPosicionActualHoja = vHoja.Index
    
    ' Solo mover si la hoja no est� ya en la posici�n correcta
    If vPosicionActualHoja <> vPosicion Then
        ' Mover la hoja a la posici�n especificada
        If vPosicion = 1 Then
            ' Si es la primera posici�n, mover antes de la primera hoja
            vHoja.Move Before:=ThisWorkbook.Worksheets(1)
        Else
            ' Obtener referencia a la hoja en la posici�n objetivo
            Set vHojaReferencia = Nothing
            On Error Resume Next
            
            If vPosicionActualHoja < vPosicion Then
                ' La hoja est� antes de su destino
                Set vHojaReferencia = ThisWorkbook.Worksheets(vPosicion - 1)
                If Not vHojaReferencia Is Nothing Then
                    vHoja.Move After:=vHojaReferencia
                End If
            Else
                ' La hoja est� despu�s de su destino
                Set vHojaReferencia = ThisWorkbook.Worksheets(vPosicion)
                If Not vHojaReferencia Is Nothing Then
                    vHoja.Move Before:=vHojaReferencia
                End If
            End If
            
            On Error GoTo ErrorHandler
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Registrar error espec�fico en log si es posible
    On Error Resume Next
    Call fun801_LogMessage("Error al mover hoja " & Chr(34) & vNombreHoja & Chr(34) & _
        " a posici�n " & CStr(vPosicion) & ": " & Err.Description, True, "", "fun803_Mover_Hoja_A_Posicion_Segura")
    On Error GoTo 0
    
End Sub

