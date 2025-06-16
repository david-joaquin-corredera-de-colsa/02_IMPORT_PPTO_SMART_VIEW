Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_06"
Option Explicit
Public Function Modificar_Scenario_Year_Entity_en_hoja_PLAH( _
    ByVal vReport_PL_AH_Name As String, _
    ByVal vFilaScenario As Integer, _
    ByVal vFilaYear As Integer, _
    ByVal vFilaEntity As Integer, _
    ByVal vColumnaInicialHeaders As Integer, _
    ByVal vColumnaFinalHeaders As Integer, _
    ByVal vScenario_xPL As String, _
    ByVal vYear_xPL As String, _
    ByVal vEntity_xPL As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: Modificar_Scenario_Year_Entity_en_hoja_PLAH
    ' Fecha y Hora de Creación: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Función para modificar las dimensiones Scenario, Year y Entity en una hoja
    ' específica de Excel, actualizando los valores en las filas correspondientes
    ' dentro de un rango de columnas determinado.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validación de parámetros de entrada
    ' 2. Verificación de existencia de la hoja objetivo
    ' 3. Obtención de referencia a la hoja de trabajo
    ' 4. Configuración del entorno para optimizar rendimiento
    ' 5. Validación de rangos de filas y columnas
    ' 6. Recorrido de columnas desde vColumnaInicialHeaders hasta vColumnaFinalHeaders
    ' 7. Asignación de valores en fila vFilaScenario con vScenario_xPL
    ' 8. Asignación de valores en fila vFilaYear con vYear_xPL
    ' 9. Asignación de valores en fila vFilaEntity con vEntity_xPL
    ' 10. Restauración del entorno de Excel
    ' 11. Registro del resultado en el sistema de logging
    '
    ' Compatibilidad: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
        
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hoja de trabajo
    Dim wsDestino As Worksheet
    
    ' Variables para bucles
    Dim i As Integer
    
    ' Variables para optimización
    Dim blnScreenUpdating As Boolean
    Dim blnEnableEvents As Boolean
    Dim xlCalculationMode As Long
    
    ' Inicialización
    strFuncion = "Modificar_Scenario_Year_Entity_en_hoja_PLAH" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "Modificar_Scenario_Year_Entity_en_hoja_PLAH"
    Modificar_Scenario_Year_Entity_en_hoja_PLAH = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validación de parámetros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando validación de parámetros de entrada...", False, "", strFuncion
    
    ' Validar nombre de hoja
    If Not fun827_ValidarNombreHoja(vReport_PL_AH_Name) Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "Nombre de hoja no válido: " & vReport_PL_AH_Name
    End If
    
    ' Validar filas
    If Not fun828_ValidarParametrosFila(vFilaScenario, vFilaYear, vFilaEntity) Then
        Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
            "Parámetros de fila no válidos. Scenario: " & vFilaScenario & _
            ", Year: " & vFilaYear & ", Entity: " & vFilaEntity
    End If
    
    ' Validar columnas
    If Not fun829_ValidarParametrosColumna(vColumnaInicialHeaders, vColumnaFinalHeaders) Then
        Err.Raise ERROR_BASE_IMPORT + 803, strFuncion, _
            "Parámetros de columna no válidos. Inicial: " & vColumnaInicialHeaders & _
            ", Final: " & vColumnaFinalHeaders
    End If
    
    ' Validar valores a asignar
    If Not fun830_ValidarValoresAsignar(vScenario_xPL, vYear_xPL, vEntity_xPL) Then
        Err.Raise ERROR_BASE_IMPORT + 804, strFuncion, _
            "Valores a asignar no válidos"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Verificación de existencia de la hoja objetivo
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Verificando existencia de hoja objetivo...", False, "", vReport_PL_AH_Name
    
    If Not fun802_SheetExists(vReport_PL_AH_Name) Then
        Err.Raise ERROR_BASE_IMPORT + 805, strFuncion, _
            "La hoja especificada no existe: " & vReport_PL_AH_Name
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Obtención de referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Obteniendo referencia a la hoja de trabajo...", False, "", vReport_PL_AH_Name
    
    Set wsDestino = ThisWorkbook.Worksheets(vReport_PL_AH_Name)
    
    If wsDestino Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 806, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vReport_PL_AH_Name
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Configuración del entorno para optimizar rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Configurando entorno para optimización...", False, "", vReport_PL_AH_Name
    
    ' Guardar configuración actual
    blnScreenUpdating = Application.ScreenUpdating
    blnEnableEvents = Application.EnableEvents
    xlCalculationMode = Application.Calculation
    
    ' Configurar para optimización
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '--------------------------------------------------------------------------
    ' 5. Validación de rangos en la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 90
    fun801_LogMessage "Validando rangos en la hoja de trabajo...", False, "", vReport_PL_AH_Name
    
    If Not fun831_ValidarRangosEnHoja(wsDestino, vFilaScenario, vFilaYear, vFilaEntity, _
                                      vColumnaInicialHeaders, vColumnaFinalHeaders) Then
        Err.Raise ERROR_BASE_IMPORT + 807, strFuncion, _
            "Los rangos especificados exceden los límites de la hoja"
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Recorrido de columnas y asignación de valores
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Iniciando recorrido de columnas para asignación de valores...", _
                      False, "", vReport_PL_AH_Name
    
    For i = vColumnaInicialHeaders To vColumnaFinalHeaders
        '----------------------------------------------------------------------
        ' 6.1. Asignación de valor Scenario en fila correspondiente
        '----------------------------------------------------------------------
        lngLineaError = 110
        wsDestino.Cells(vFilaScenario, i).Value = vScenario_xPL
        
        '----------------------------------------------------------------------
        ' 6.2. Asignación de valor Year en fila correspondiente
        '----------------------------------------------------------------------
        lngLineaError = 120
        wsDestino.Cells(vFilaYear, i).Value = vYear_xPL
        
        '----------------------------------------------------------------------
        ' 6.3. Asignación de valor Entity en fila correspondiente
        '----------------------------------------------------------------------
        lngLineaError = 130
        wsDestino.Cells(vFilaEntity, i).Value = vEntity_xPL
    Next i
    
    '--------------------------------------------------------------------------
    ' 7. Restauración del entorno de Excel
    '--------------------------------------------------------------------------
    lngLineaError = 140
    fun801_LogMessage "Restaurando configuración del entorno...", False, "", vReport_PL_AH_Name
    
    Application.Calculation = xlCalculationMode
    Application.EnableEvents = blnEnableEvents
    Application.ScreenUpdating = blnScreenUpdating
    
    '--------------------------------------------------------------------------
    ' 8. Registro del resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 150
    fun801_LogMessage "Modificación completada exitosamente. Columnas procesadas: " & _
                      (vColumnaFinalHeaders - vColumnaInicialHeaders + 1), _
                      False, "", vReport_PL_AH_Name
    
    Modificar_Scenario_Year_Entity_en_hoja_PLAH = True
    Exit Function

GestorErrores:
    ' Restaurar configuración del entorno en caso de error
    On Error Resume Next
    Application.Calculation = xlCalculationMode
    Application.EnableEvents = blnEnableEvents
    Application.ScreenUpdating = blnScreenUpdating
    On Error GoTo 0
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja: " & vReport_PL_AH_Name & vbCrLf & _
                      "Parámetros: Scenario(" & vFilaScenario & "), Year(" & vFilaYear & _
                      "), Entity(" & vFilaEntity & "), Cols(" & vColumnaInicialHeaders & _
                      "-" & vColumnaFinalHeaders & ")"
    
    fun801_LogMessage strMensajeError, True, "", vReport_PL_AH_Name
    Modificar_Scenario_Year_Entity_en_hoja_PLAH = False
End Function

Public Function fun824_LimpiarFilasExcedentes(ByRef ws As Worksheet, _
                                             ByVal vFila_Final_Limite As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun824_LimpiarFilasExcedentes
    ' Fecha y Hora de Creación: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Limpia todas las filas que estén por encima del límite especificado
    ' Borra tanto contenido como formatos para optimizar el archivo
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde limpiar
    ' - vFila_Final_Limite: Número de fila límite (se borran filas superiores)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngUltimaFilaConDatos As Long
    
    ' Validar parámetros
    If ws Is Nothing Then
        fun824_LimpiarFilasExcedentes = False
        Exit Function
    End If
    
    If vFila_Final_Limite < 1 Then
        fun824_LimpiarFilasExcedentes = False
        Exit Function
    End If
    
    ' Obtener última fila con datos (método compatible Excel 97-365)
    lngUltimaFilaConDatos = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Si hay filas excedentes, limpiarlas completamente
    If lngUltimaFilaConDatos > vFila_Final_Limite Then
        Application.ScreenUpdating = False
        
        ' Limpiar contenido y formatos (compatible Excel 97-365)
        ws.Range(ws.Cells(vFila_Final_Limite + 1, 1), _
                 ws.Cells(lngUltimaFilaConDatos, ws.Columns.Count)).Clear
        
        Application.ScreenUpdating = True
    End If
    
    fun824_LimpiarFilasExcedentes = True
    Exit Function
    
GestorErrores:
    Application.ScreenUpdating = True
    fun824_LimpiarFilasExcedentes = False
End Function

Public Function fun825_LimpiarColumnasExcedentes(ByRef ws As Worksheet, _
                                                ByVal vColumna_Final_Limite As Long) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun825_LimpiarColumnasExcedentes
    ' Fecha y Hora de Creación: 2025-06-03 00:18:41 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Limpia todas las columnas que estén por encima del límite especificado
    ' Borra tanto contenido como formatos para optimizar el archivo
    '
    ' Parámetros:
    ' - ws: Hoja de trabajo donde limpiar
    ' - vColumna_Final_Limite: Número de columna límite (se borran columnas superiores)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim lngUltimaColumnaConDatos As Long
    
    ' Validar parámetros
    If ws Is Nothing Then
        fun825_LimpiarColumnasExcedentes = False
        Exit Function
    End If
    
    If vColumna_Final_Limite < 1 Then
        fun825_LimpiarColumnasExcedentes = False
        Exit Function
    End If
    
    ' Obtener última columna con datos (método compatible Excel 97-365)
    lngUltimaColumnaConDatos = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Si hay columnas excedentes, limpiarlas completamente
    If lngUltimaColumnaConDatos > vColumna_Final_Limite Then
        Application.ScreenUpdating = False
        
        ' Limpiar contenido y formatos (compatible Excel 97-365)
        ws.Range(ws.Cells(1, vColumna_Final_Limite + 1), _
                 ws.Cells(ws.Rows.Count, lngUltimaColumnaConDatos)).Clear
        
        Application.ScreenUpdating = True
    End If
    
    fun825_LimpiarColumnasExcedentes = True
    Exit Function
    
GestorErrores:
    Application.ScreenUpdating = True
    fun825_LimpiarColumnasExcedentes = False
End Function

Public Function fun826_ConfigurarPalabrasClave(Optional ByVal strPrimeraFila As String = "", _
                                              Optional ByVal strPrimeraColumna As String = "", _
                                              Optional ByVal strUltimaFila As String = "", _
                                              Optional ByVal strUltimaColumna As String = "") As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun826_ConfigurarPalabrasClave
    ' Fecha y Hora de Creación: 2025-06-03 03:19:45 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Permite configurar las palabras clave utilizadas para detectar rangos
    ' de datos en las hojas de trabajo.
    '
    ' Parámetros (todos opcionales):
    ' - strPrimeraFila: Palabra clave para buscar primera fila
    ' - strPrimeraColumna: Palabra clave para buscar primera columna
    ' - strUltimaFila: Palabra clave para buscar última fila
    ' - strUltimaColumna: Palabra clave para buscar última columna
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    ' Solo actualizar las variables que se proporcionen
    If Len(Trim(strPrimeraFila)) > 0 Then
        vPalabraClave_PrimeraFila = Trim(strPrimeraFila)
    End If
    
    If Len(Trim(strPrimeraColumna)) > 0 Then
        vPalabraClave_PrimeraColumna = Trim(strPrimeraColumna)
    End If
    
    If Len(Trim(strUltimaFila)) > 0 Then
        vPalabraClave_UltimaFila = Trim(strUltimaFila)
    End If
    
    If Len(Trim(strUltimaColumna)) > 0 Then
        vPalabraClave_UltimaColumna = Trim(strUltimaColumna)
    End If
    
    fun826_ConfigurarPalabrasClave = True
    Exit Function
    
GestorErrores:
    fun826_ConfigurarPalabrasClave = False
End Function

Public Function fun823_OcultarHojaSiVisible(ByRef ws As Worksheet) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_OcultarHojaSiVisible
    ' Fecha y Hora de Creación: 2025-06-03 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Oculta una hoja si está visible
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    If ws.Visible = xlSheetVisible Then
        ws.Visible = xlSheetHidden
        fun823_OcultarHojaSiVisible = True
    Else
        fun823_OcultarHojaSiVisible = False
    End If
    
    Exit Function
    
GestorErrores:
    fun823_OcultarHojaSiVisible = False
End Function

Public Function fun823_MostrarHojaSiOculta(vNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_MostrarHojaSiOculta
    ' Fecha y Hora de Creación: 2025-06-08 04:25:04 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Muestra una hoja si está oculta
    '******************************************************************************
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(vNombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then
        fun823_MostrarHojaSiOculta = False
    Else
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
        End If
        fun823_MostrarHojaSiOculta = True
    End If
End Function



Public Function fun821_ComenzarPorPrefijo(ByVal strTexto As String, ByVal strPrefijo As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun821_ComenzarPorPrefijo
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    If Len(strTexto) >= Len(strPrefijo) Then
        fun821_ComenzarPorPrefijo = (Left(strTexto, Len(strPrefijo)) = strPrefijo)
    Else
        fun821_ComenzarPorPrefijo = False
    End If
    Exit Function
    
ErrorHandler:
    fun821_ComenzarPorPrefijo = False
End Function

Public Function fun822_ValidarFormatoSufijoFecha(ByVal strNombreHoja As String, _
                                                ByVal strPrefijo As String, _
                                                ByVal intLongitudSufijo As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun822_ValidarFormatoSufijoFecha
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    Dim intLongitudEsperada As Integer
    intLongitudEsperada = Len(strPrefijo) + intLongitudSufijo
    
    ' Validar longitud total
    If Len(strNombreHoja) = intLongitudEsperada Then
        fun822_ValidarFormatoSufijoFecha = True
    Else
        fun822_ValidarFormatoSufijoFecha = False
    End If
    Exit Function
    
ErrorHandler:
    fun822_ValidarFormatoSufijoFecha = False
End Function

Public Function fun823_ExtraerSufijoFecha(ByVal strNombreHoja As String, _
                                         ByVal intLongitudSufijo As Integer) As String
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun823_ExtraerSufijoFecha
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    If Len(strNombreHoja) >= intLongitudSufijo Then
        fun823_ExtraerSufijoFecha = Right(strNombreHoja, intLongitudSufijo)
    Else
        fun823_ExtraerSufijoFecha = ""
    End If
    Exit Function
    
ErrorHandler:
    fun823_ExtraerSufijoFecha = ""
End Function

Public Function fun824_CompararSufijosFecha(ByVal strSufijo1 As String, _
                                           ByVal strSufijo2 As String) As Integer
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun824_CompararSufijosFecha
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    ' Retorna: >0 si strSufijo1 > strSufijo2, 0 si iguales, <0 si strSufijo1 < strSufijo2
    '******************************************************************************
    On Error GoTo ErrorHandler
    
    If strSufijo2 = "" Then
        fun824_CompararSufijosFecha = 1  ' strSufijo1 es mayor
    ElseIf strSufijo1 > strSufijo2 Then
        fun824_CompararSufijosFecha = 1
    ElseIf strSufijo1 < strSufijo2 Then
        fun824_CompararSufijosFecha = -1
    Else
        fun824_CompararSufijosFecha = 0
    End If
    Exit Function
    
ErrorHandler:
    fun824_CompararSufijosFecha = 0
End Function

Public Function fun825_CopiarHojaConNuevoNombre(ByVal strHojaOrigen As String, _
                                               ByVal strHojaDestino As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun825_CopiarHojaConNuevoNombre
    ' Fecha y Hora de Creación: 2025-06-03 06:00:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Crea una copia completa de una hoja de trabajo existente y le asigna
    ' un nuevo nombre. Maneja conflictos de nombres eliminando hojas existentes.
    '
    ' Pasos:
    ' 1. Validar que la hoja origen existe
    ' 2. Generar nombre de destino si no se proporciona
    ' 3. Eliminar hoja destino si ya existe
    ' 4. Copiar hoja origen con nuevo nombre
    ' 5. Verificar que la copia se creó correctamente
    '
    ' Parámetros:
    ' - strHojaOrigen: Nombre de la hoja a copiar
    ' - strHojaDestino: Nombre para la nueva hoja copiada
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim strNombreDestino As String
    
    ' Inicialización
    strFuncion = "fun825_CopiarHojaConNuevoNombre" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "fun825_CopiarHojaConNuevoNombre"
    fun825_CopiarHojaConNuevoNombre = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que la hoja origen existe
    '--------------------------------------------------------------------------
    lngLineaError = 30
    If Len(Trim(strHojaOrigen)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 851, strFuncion, _
            "El nombre de la hoja origen está vacío"
    End If
    
    If Not fun802_SheetExists(strHojaOrigen) Then
        Err.Raise ERROR_BASE_IMPORT + 852, strFuncion, _
            "La hoja origen no existe: " & strHojaOrigen
    End If
    
    Set wsOrigen = ThisWorkbook.Worksheets(strHojaOrigen)
    
    '--------------------------------------------------------------------------
    ' 2. Preparar nombre de destino
    '--------------------------------------------------------------------------
    lngLineaError = 40
    If Len(Trim(strHojaDestino)) = 0 Then
        ' Generar nombre automático basado en timestamp
        strNombreDestino = strHojaOrigen & "_Copia_" & Format(Now(), "yyyymmdd_hhmmss")
    Else
        strNombreDestino = Trim(strHojaDestino)
    End If
    
    ' Validar longitud del nombre (Excel tiene límite de 31 caracteres)
    If Len(strNombreDestino) > 31 Then
        strNombreDestino = Left(strNombreDestino, 31)
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Eliminar hoja destino si ya existe
    '--------------------------------------------------------------------------
    lngLineaError = 50
    If fun802_SheetExists(strNombreDestino) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(strNombreDestino).Delete
        Application.DisplayAlerts = True
        
        fun801_LogMessage "Hoja existente eliminada: " & strNombreDestino, False, "", strFuncion
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Copiar hoja origen con nuevo nombre
    '--------------------------------------------------------------------------
    lngLineaError = 60
    Application.ScreenUpdating = False
    
    ' Copiar la hoja al final del libro
    wsOrigen.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ' Obtener referencia a la hoja recién copiada
    Set wsDestino = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ' Asignar nuevo nombre
    wsDestino.Name = strNombreDestino
    
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 5. Verificar que la copia se creó correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 70
    If Not fun802_SheetExists(strNombreDestino) Then
        Err.Raise ERROR_BASE_IMPORT + 853, strFuncion, _
            "Error al verificar la creación de la hoja copiada: " & strNombreDestino
    End If
    
    fun801_LogMessage "Hoja copiada exitosamente: " & strHojaOrigen & " ? " & strNombreDestino, _
                      False, "", strFuncion
    
    fun825_CopiarHojaConNuevoNombre = True
    Exit Function

GestorErrores:
    ' Restaurar configuración
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja Origen: " & strHojaOrigen & vbCrLf & _
                      "Hoja Destino: " & strHojaDestino
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun825_CopiarHojaConNuevoNombre = False
End Function

