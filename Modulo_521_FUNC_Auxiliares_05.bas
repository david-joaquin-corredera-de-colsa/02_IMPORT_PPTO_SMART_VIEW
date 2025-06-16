Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_05"

Option Explicit


'******************************************************************************
' FUNCIONES AUXILIARES PARA VALIDACIÓN
'******************************************************************************

Public Function fun827_ValidarNombreHoja(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun827_ValidarNombreHoja
    ' Fecha y Hora de Creación: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Valida que el nombre de hoja sea válido y no esté vacío
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun827_ValidarNombreHoja = False
    
    ' Verificar que no esté vacío
    If Len(Trim(strNombreHoja)) = 0 Then
        Exit Function
    End If
    
    ' Verificar que no contenga caracteres no válidos para nombres de hoja
    If InStr(strNombreHoja, "[") > 0 Or InStr(strNombreHoja, "]") > 0 Or _
       InStr(strNombreHoja, ":") > 0 Or InStr(strNombreHoja, "*") > 0 Or _
       InStr(strNombreHoja, "?") > 0 Or InStr(strNombreHoja, "/") > 0 Or _
       InStr(strNombreHoja, "\") > 0 Then
        Exit Function
    End If
    
    ' Verificar longitud máxima (31 caracteres para Excel)
    If Len(strNombreHoja) > 31 Then
        Exit Function
    End If
    
    fun827_ValidarNombreHoja = True
    Exit Function
    
ErrorHandler:
    fun827_ValidarNombreHoja = False
End Function

Public Function fun828_ValidarParametrosFila(ByVal vFilaScenario As Integer, _
                                            ByVal vFilaYear As Integer, _
                                            ByVal vFilaEntity As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun828_ValidarParametrosFila
    ' Fecha y Hora de Creación: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Valida que los parámetros de fila sean válidos
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun828_ValidarParametrosFila = False
    
    ' Verificar que sean valores positivos
    If vFilaScenario <= 0 Or vFilaYear <= 0 Or vFilaEntity <= 0 Then
        Exit Function
    End If
    
    ' Verificar que no excedan el límite máximo de Excel (compatible con Excel 97)
    If vFilaScenario > 65536 Or vFilaYear > 65536 Or vFilaEntity > 65536 Then
        Exit Function
    End If
    
    fun828_ValidarParametrosFila = True
    Exit Function
    
ErrorHandler:
    fun828_ValidarParametrosFila = False
End Function

Public Function fun829_ValidarParametrosColumna(ByVal vColumnaInicial As Integer, _
                                               ByVal vColumnaFinal As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun829_ValidarParametrosColumna
    ' Fecha y Hora de Creación: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Valida que los parámetros de columna sean válidos
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun829_ValidarParametrosColumna = False
    
    ' Verificar que sean valores positivos
    If vColumnaInicial <= 0 Or vColumnaFinal <= 0 Then
        Exit Function
    End If
    
    ' Verificar que la columna inicial sea menor o igual que la final
    If vColumnaInicial > vColumnaFinal Then
        Exit Function
    End If
    
    ' Verificar que no excedan el límite máximo de Excel (compatible con Excel 97: 256 columnas)
    If vColumnaInicial > 256 Or vColumnaFinal > 256 Then
        Exit Function
    End If
    
    fun829_ValidarParametrosColumna = True
    Exit Function
    
ErrorHandler:
    fun829_ValidarParametrosColumna = False
End Function

Public Function fun830_ValidarValoresAsignar(ByVal vScenario As String, _
                                            ByVal vYear As String, _
                                            ByVal vEntity As String) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun830_ValidarValoresAsignar
    ' Fecha y Hora de Creación: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Valida que los valores a asignar sean válidos (pueden estar vacíos)
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    ' En este caso, permitimos valores vacíos ya que podrían ser válidos
    ' Solo verificamos que no sean Nothing (aunque al ser String esto no aplica)
    
    ' Verificar longitud máxima razonable para evitar problemas de memoria
    If Len(vScenario) > 255 Or Len(vYear) > 255 Or Len(vEntity) > 255 Then
        fun830_ValidarValoresAsignar = False
        Exit Function
    End If
    
    fun830_ValidarValoresAsignar = True
    Exit Function
    
ErrorHandler:
    fun830_ValidarValoresAsignar = False
End Function

Public Function fun831_ValidarRangosEnHoja(ByRef ws As Worksheet, _
                                          ByVal vFilaScenario As Integer, _
                                          ByVal vFilaYear As Integer, _
                                          ByVal vFilaEntity As Integer, _
                                          ByVal vColumnaInicial As Integer, _
                                          ByVal vColumnaFinal As Integer) As Boolean
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun831_ValidarRangosEnHoja
    ' Fecha y Hora de Creación: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción: Valida que los rangos especificados existan en la hoja
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun831_ValidarRangosEnHoja = False
    
    ' Verificar que la hoja sea válida
    If ws Is Nothing Then
        Exit Function
    End If
    
    ' Verificar que las filas estén dentro del rango de la hoja
    If vFilaScenario > ws.Rows.Count Or vFilaYear > ws.Rows.Count Or vFilaEntity > ws.Rows.Count Then
        Exit Function
    End If
    
    ' Verificar que las columnas estén dentro del rango de la hoja
    If vColumnaInicial > ws.Columns.Count Or vColumnaFinal > ws.Columns.Count Then
        Exit Function
    End If
    
    ' Intentar acceder a las celdas para verificar que son accesibles
    On Error Resume Next
    Dim testValue As Variant
    testValue = ws.Cells(vFilaScenario, vColumnaInicial).Value
    testValue = ws.Cells(vFilaYear, vColumnaFinal).Value
    testValue = ws.Cells(vFilaEntity, vColumnaInicial).Value
    
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    fun831_ValidarRangosEnHoja = True
    Exit Function
    
ErrorHandler:
    fun831_ValidarRangosEnHoja = False
End Function

Public Function Convertir_RangoCellsCells_a_RangoCFCF(ByVal vFilaInicial As Integer, _
                                                      ByVal vFilaFinal As Integer, _
                                                      ByVal vColumnaInicial As Integer, _
                                                      ByVal vColumnaFinal As Integer) As String
    
    '******************************************************************************
    ' FUNCIÓN: Convertir_RangoCellsCells_a_RangoCFCF
    ' FECHA Y HORA DE CREACIÓN: 2025-06-15 11:29:40 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCIÓN:
    ' Convierte coordenadas numéricas de filas y columnas a formato de rango de Excel
    ' estándar tipo "A5:P100". Función auxiliar para generación dinámica de rangos
    ' de celdas en operaciones de manipulación de hojas de cálculo.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables de control de errores y validación
    ' 2. Validación exhaustiva de parámetros de entrada (rangos válidos)
    ' 3. Verificación de lógica de coordenadas (inicial <= final)
    ' 4. Conversión de números de columna a letras usando función auxiliar
    ' 5. Construcción del string de rango en formato Excel estándar
    ' 6. Validación del resultado generado antes del retorno
    ' 7. Logging de operación para debugging y auditoría
    ' 8. Retorno del string de rango formateado
    ' 9. Manejo exhaustivo de errores con información detallada
    ' 10. Limpieza de recursos y logging de errores en caso de fallo
    '
    ' PARÁMETROS:
    ' - vFilaInicial (Integer): Número de fila inicial (debe ser >= 1)
    ' - vFilaFinal (Integer): Número de fila final (debe ser >= vFilaInicial)
    ' - vColumnaInicial (Integer): Número de columna inicial (debe ser >= 1)
    ' - vColumnaFinal (Integer): Número de columna final (debe ser >= vColumnaInicial)
    '
    ' RETORNA: String - Rango en formato Excel (ej: "A5:P100") o cadena vacía si error
    '
    ' EJEMPLOS DE USO:
    ' Dim strRango As String
    ' strRango = Convertir_RangoCellsCells_a_RangoCFCF(5, 100, 1, 16)    ' Devuelve "A5:P100"
    ' strRango = Convertir_RangoCellsCells_a_RangoCFCF(1, 1, 1, 1)       ' Devuelve "A1:A1"
    ' strRango = Convertir_RangoCellsCells_a_RangoCFCF(10, 20, 5, 8)     ' Devuelve "E10:H20"
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento
    Dim strColumnaInicialLetra As String
    Dim strColumnaFinalLetra As String
    Dim strRangoResultado As String
    
    ' Inicialización
    strFuncion = "Convertir_RangoCellsCells_a_RangoCFCF"
    Convertir_RangoCellsCells_a_RangoCFCF = ""
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables de control de errores y validación
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    ' Inicializar variables de trabajo
    strColumnaInicialLetra = ""
    strColumnaFinalLetra = ""
    strRangoResultado = ""
    
    '--------------------------------------------------------------------------
    ' 2. Validación exhaustiva de parámetros de entrada
    '--------------------------------------------------------------------------
    lngLineaError = 40
    
    ' Validar fila inicial
    If vFilaInicial < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9101, strFuncion, _
            "Fila inicial debe ser mayor que 0. Valor recibido: " & vFilaInicial
    End If
    
    ' Validar fila final
    If vFilaFinal < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9102, strFuncion, _
            "Fila final debe ser mayor que 0. Valor recibido: " & vFilaFinal
    End If
    
    ' Validar columna inicial
    If vColumnaInicial < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9103, strFuncion, _
            "Columna inicial debe ser mayor que 0. Valor recibido: " & vColumnaInicial
    End If
    
    ' Validar columna final
    If vColumnaFinal < 1 Then
        Err.Raise ERROR_BASE_IMPORT + 9104, strFuncion, _
            "Columna final debe ser mayor que 0. Valor recibido: " & vColumnaFinal
    End If
    
    ' Validar límites máximos de Excel (compatible con Excel 97-365)
    If vFilaInicial > 65536 Then
        Err.Raise ERROR_BASE_IMPORT + 9105, strFuncion, _
            "Fila inicial excede límite máximo de Excel (65536). Valor recibido: " & vFilaInicial
    End If
    
    If vFilaFinal > 65536 Then
        Err.Raise ERROR_BASE_IMPORT + 9106, strFuncion, _
            "Fila final excede límite máximo de Excel (65536). Valor recibido: " & vFilaFinal
    End If
    
    If vColumnaInicial > 256 Then
        Err.Raise ERROR_BASE_IMPORT + 9107, strFuncion, _
            "Columna inicial excede límite máximo de Excel (256). Valor recibido: " & vColumnaInicial
    End If
    
    If vColumnaFinal > 256 Then
        Err.Raise ERROR_BASE_IMPORT + 9108, strFuncion, _
            "Columna final excede límite máximo de Excel (256). Valor recibido: " & vColumnaFinal
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Verificación de lógica de coordenadas
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Verificar que fila inicial <= fila final
    If vFilaInicial > vFilaFinal Then
        Err.Raise ERROR_BASE_IMPORT + 9109, strFuncion, _
            "Fila inicial (" & vFilaInicial & ") debe ser menor o igual que fila final (" & vFilaFinal & ")"
    End If
    
    ' Verificar que columna inicial <= columna final
    If vColumnaInicial > vColumnaFinal Then
        Err.Raise ERROR_BASE_IMPORT + 9110, strFuncion, _
            "Columna inicial (" & vColumnaInicial & ") debe ser menor o igual que columna final (" & vColumnaFinal & ")"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Conversión de números de columna a letras
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Convertir columna inicial a letra usando función auxiliar
    strColumnaInicialLetra = fun801_ConvertirNumeroColumnaALetra(vColumnaInicial)
    
    If Len(strColumnaInicialLetra) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9111, strFuncion, _
            "Error al convertir columna inicial a letra. Columna: " & vColumnaInicial
    End If
    
    ' Convertir columna final a letra usando función auxiliar
    strColumnaFinalLetra = fun801_ConvertirNumeroColumnaALetra(vColumnaFinal)
    
    If Len(strColumnaFinalLetra) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9112, strFuncion, _
            "Error al convertir columna final a letra. Columna: " & vColumnaFinal
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Construcción del string de rango en formato Excel estándar
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Construir el rango en formato "COLUMNA_INICIAL+FILA_INICIAL:COLUMNA_FINAL+FILA_FINAL"
    strRangoResultado = strColumnaInicialLetra & CStr(vFilaInicial) & Chr(58) & _
                        strColumnaFinalLetra & CStr(vFilaFinal)
    
    '--------------------------------------------------------------------------
    ' 6. Validación del resultado generado
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Verificar que el resultado no esté vacío
    If Len(strRangoResultado) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9113, strFuncion, _
            "Error al generar string de rango - resultado vacío"
    End If
    
    ' Verificar que contiene el separador de rango (:)
    If InStr(strRangoResultado, Chr(58)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9114, strFuncion, _
            "Error en formato de rango - separador no encontrado: " & strRangoResultado
    End If
    
    ' Verificar longitud mínima (ej: "A1:A1" = 5 caracteres mínimo)
    If Len(strRangoResultado) < 5 Then
        Err.Raise ERROR_BASE_IMPORT + 9115, strFuncion, _
            "Longitud de rango inválida: " & strRangoResultado & " (Longitud: " & Len(strRangoResultado) & ")"
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Logging de operación para debugging y auditoría
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    Call fun801_LogMessage("CONVERSIÓN EXITOSA - Rango generado: " & Chr(34) & strRangoResultado & Chr(34) & _
        " desde coordenadas F(" & vFilaInicial & ":" & vFilaFinal & ") C(" & _
        vColumnaInicial & ":" & vColumnaFinal & ") = (" & strColumnaInicialLetra & ":" & _
        strColumnaFinalLetra & ")", False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 8. Retorno del string de rango formateado
    '--------------------------------------------------------------------------
    lngLineaError = 100
    Convertir_RangoCellsCells_a_RangoCFCF = strRangoResultado
    
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 9. Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Parámetros de entrada:" & vbCrLf & _
                      "  - Fila inicial: " & vFilaInicial & vbCrLf & _
                      "  - Fila final: " & vFilaFinal & vbCrLf & _
                      "  - Columna inicial: " & vColumnaInicial & vbCrLf & _
                      "  - Columna final: " & vColumnaFinal & vbCrLf & _
                      "Variables de trabajo:" & vbCrLf & _
                      "  - Columna inicial letra: " & Chr(34) & strColumnaInicialLetra & Chr(34) & vbCrLf & _
                      "  - Columna final letra: " & Chr(34) & strColumnaFinalLetra & Chr(34) & vbCrLf & _
                      "  - Rango resultado: " & Chr(34) & strRangoResultado & Chr(34) & vbCrLf & _
                      "Fecha y Hora: " & Now() & vbCrLf & _
                      "Compatibilidad: Excel 97/2003/2007/365, OneDrive/SharePoint/Teams"
    
    '--------------------------------------------------------------------------
    ' 10. Logging de errores y limpieza de recursos
    '--------------------------------------------------------------------------
    
    ' Registrar error completo en log del sistema
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Para debugging en desarrollo
    Debug.Print strMensajeError
    
    ' Retornar cadena vacía para indicar error
    Convertir_RangoCellsCells_a_RangoCFCF = ""
End Function

Public Function fun801_ConvertirNumeroColumnaALetra(ByVal vNumeroColumna As Integer) As String
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun801_ConvertirNumeroColumnaALetra
    ' FECHA Y HORA DE CREACIÓN: 2025-06-15 11:29:40 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCIÓN:
    ' Convierte un número de columna (1, 2, 3...) a su letra correspondiente
    ' en Excel (A, B, C, AA, AB...). Función auxiliar para conversión de rangos.
    '
    ' PARÁMETROS:
    ' - vNumeroColumna (Integer): Número de columna (1-256 para compatibilidad Excel 97)
    '
    ' RETORNA: String - Letra(s) de columna Excel o cadena vacía si error
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim strResultado As String
    Dim intNumero As Integer
    
    ' Inicialización
    fun801_ConvertirNumeroColumnaALetra = ""
    
    ' Validar parámetro
    If vNumeroColumna < 1 Or vNumeroColumna > 256 Then
        Exit Function
    End If
    
    ' Algoritmo de conversión a base 26 (letras A-Z)
    intNumero = vNumeroColumna
    strResultado = ""
    
    Do While intNumero > 0
        intNumero = intNumero - 1  ' Ajustar para base 0
        strResultado = Chr(65 + (intNumero Mod 26)) & strResultado
        intNumero = intNumero \ 26
    Loop
    
    fun801_ConvertirNumeroColumnaALetra = strResultado
    Exit Function
    
ErrorHandler:
    fun801_ConvertirNumeroColumnaALetra = ""
End Function

Public Function Contiene_Scenario_Year_Entity(ByVal vSheet As String, _
                                             ByVal vEscenario As String, _
                                             ByVal vAnio As String, _
                                             ByVal vSociedad As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN: Contiene_Scenario_Year_Entity
    ' FECHA Y HORA DE CREACIÓN: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCIÓN:
    ' Función que verifica si una hoja específica contiene tres valores exactos:
    ' escenario, año y sociedad. Realiza búsqueda exhaustiva en toda la hoja
    ' para determinar si los tres valores están presentes como contenido completo
    ' de celdas individuales.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicialización de variables de control de errores y optimización
    ' 2. Validación de parámetros de entrada y longitudes
    ' 3. Configuración de optimizaciones de rendimiento (pantalla, cálculo)
    ' 4. Verificación de existencia de la hoja especificada
    ' 5. Obtención de referencia a la hoja de trabajo
    ' 6. Determinación del rango usado para búsqueda eficiente
    ' 7. Búsqueda del primer valor (escenario) con coincidencia exacta
    ' 8. Búsqueda del segundo valor (año) con coincidencia exacta
    ' 9. Búsqueda del tercer valor (sociedad) con coincidencia exacta
    ' 10. Evaluación de resultados y determinación del valor de retorno
    ' 11. Registro de resultados en log del sistema
    ' 12. Restauración de configuraciones de optimización
    ' 13. Manejo exhaustivo de errores con información detallada
    '
    ' PARÁMETROS:
    ' - vSheet (String): Nombre de la hoja donde realizar la búsqueda
    ' - vEscenario (String): Valor del escenario a buscar (coincidencia exacta)
    ' - vAnio (String): Valor del año a buscar (coincidencia exacta)
    ' - vSociedad (String): Valor de la sociedad a buscar (coincidencia exacta)
    '
    ' VALOR DE RETORNO:
    ' - Boolean: True si los tres valores existen en la hoja, False en caso contrario
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    ' VERSIÓN: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimización
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas y búsqueda
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim blnHojaExiste As Boolean
    Dim rngUsedRange As Range
    
    ' Variables para resultados de búsqueda
    Dim blnExisteEscenario As Boolean
    Dim blnExisteAnio As Boolean
    Dim blnExisteSociedad As Boolean
    
    ' Variables para log de resultados
    Dim strMensajeLog As String
    
    ' Inicialización
    strFuncion = "Contiene_Scenario_Year_Entity"
    Contiene_Scenario_Year_Entity = False
    lngLineaError = 0
    blnExisteEscenario = False
    blnExisteAnio = False
    blnExisteSociedad = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicialización de variables de control de errores y optimización
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando búsqueda de valores en hoja", False, "", vSheet
    
    ' Almacenar configuraciones originales para restaurar después
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Validación de parámetros de entrada y longitudes
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Validar que el nombre de la hoja no esté vacío
    If Len(Trim(vSheet)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "Parámetro vSheet está vacío"
    End If
    
    ' Validar que el escenario no esté vacío
    If Len(Trim(vEscenario)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
            "Parámetro vEscenario está vacío"
    End If
    
    ' Validar que el año no esté vacío
    If Len(Trim(vAnio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 803, strFuncion, _
            "Parámetro vAnio está vacío"
    End If
    
    ' Validar que la sociedad no esté vacía
    If Len(Trim(vSociedad)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 804, strFuncion, _
            "Parámetro vSociedad está vacío"
    End If
    
    ' Validar longitudes máximas razonables (compatibilidad Excel 97-365)
    If Len(Trim(vSheet)) > 31 Then
        Err.Raise ERROR_BASE_IMPORT + 805, strFuncion, _
            "Nombre de hoja demasiado largo: " & Len(Trim(vSheet)) & " caracteres"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Configuración de optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Desactivar actualización de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar cálculo automático para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 4. Verificación de existencia de la hoja especificada
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Obtener referencia al libro actual
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    ' Verificar que tenemos una referencia válida al libro
    If wb Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 806, strFuncion, _
            "No se pudo obtener referencia al libro de trabajo"
    End If
    
    ' Verificar existencia de la hoja usando función auxiliar existente del proyecto
    blnHojaExiste = fun801_VerificarExistenciaHoja(wb, vSheet)
    
    If Not blnHojaExiste Then
        Err.Raise ERROR_BASE_IMPORT + 807, strFuncion, _
            "La hoja especificada no existe: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Obtención de referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 90
    Set ws = wb.Worksheets(vSheet)
    
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 808, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Determinación del rango usado para búsqueda eficiente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Obtener rango usado de la hoja para optimizar búsqueda
    Set rngUsedRange = ws.UsedRange
    
    ' Verificar que la hoja tiene contenido
    If rngUsedRange Is Nothing Then
        fun801_LogMessage "Hoja está vacía, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    ' Verificar que el rango usado no está vacío
    If rngUsedRange.Cells.Count = 0 Then
        fun801_LogMessage "Rango usado está vacío, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Búsqueda del primer valor (escenario) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 110
    
    blnExisteEscenario = fun801_BuscarValorExactoEnRango(rngUsedRange, vEscenario)
    
    fun801_LogMessage "Búsqueda escenario " & Chr(34) & vEscenario & Chr(34) & _
        " resultado: " & blnExisteEscenario, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 8. Búsqueda del segundo valor (año) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
    blnExisteAnio = fun801_BuscarValorExactoEnRango(rngUsedRange, vAnio)
    
    fun801_LogMessage "Búsqueda año " & Chr(34) & vAnio & Chr(34) & _
        " resultado: " & blnExisteAnio, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 9. Búsqueda del tercer valor (sociedad) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 130
    
    blnExisteSociedad = fun801_BuscarValorExactoEnRango(rngUsedRange, vSociedad)
    
    fun801_LogMessage "Búsqueda sociedad " & Chr(34) & vSociedad & Chr(34) & _
        " resultado: " & blnExisteSociedad, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 10. Evaluación de resultados y determinación del valor de retorno
    '--------------------------------------------------------------------------
    lngLineaError = 140
    
    ' La función retorna True solo si los tres valores existen
    If blnExisteEscenario And blnExisteAnio And blnExisteSociedad Then
        Contiene_Scenario_Year_Entity = True
        strMensajeLog = "ÉXITO - Los tres valores existen en la hoja"
    Else
        Contiene_Scenario_Year_Entity = False
        strMensajeLog = "RESULTADO - Valores faltantes: "
        If Not blnExisteEscenario Then strMensajeLog = strMensajeLog & "Escenario "
        If Not blnExisteAnio Then strMensajeLog = strMensajeLog & "Año "
        If Not blnExisteSociedad Then strMensajeLog = strMensajeLog & "Sociedad "
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Registro de resultados en log del sistema
    '--------------------------------------------------------------------------
    lngLineaError = 150
    
    fun801_LogMessage strMensajeLog & " - Hoja: " & vSheet & _
        ", Escenario: " & Chr(34) & vEscenario & Chr(34) & _
        ", Año: " & Chr(34) & vAnio & Chr(34) & _
        ", Sociedad: " & Chr(34) & vSociedad & Chr(34) & _
        ", Resultado final: " & Contiene_Scenario_Year_Entity, _
        False, "", vSheet

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 12. Restauración de configuraciones de optimización
    '--------------------------------------------------------------------------
    lngLineaError = 160
    
    ' Restaurar configuración original de actualización de pantalla
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Restaurar configuración original de cálculo
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    ' Restaurar configuración original de eventos
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set rngUsedRange = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    fun801_LogMessage "Búsqueda completada exitosamente", False, "", vSheet
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 13. Manejo exhaustivo de errores con información detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja: " & vSheet & vbCrLf & _
                      "Escenario: " & Chr(34) & vEscenario & Chr(34) & vbCrLf & _
                      "Año: " & Chr(34) & vAnio & Chr(34) & vbCrLf & _
                      "Sociedad: " & Chr(34) & vSociedad & Chr(34) & vbCrLf & _
                      "Fecha y Hora: " & Now()
    
    ' Registrar error en log del sistema
    fun801_LogMessage strMensajeError, True, "", vSheet
    
    ' Log del error para debugging
    Debug.Print strMensajeError
    
    ' Restaurar configuraciones en caso de error
    On Error Resume Next
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set rngUsedRange = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    ' Retornar False para indicar error
    Contiene_Scenario_Year_Entity = False
End Function

Public Function fun801_BuscarValorExactoEnRango(ByRef rngBusqueda As Range, _
                                               ByVal strValorBuscado As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN AUXILIAR: fun801_BuscarValorExactoEnRango
    ' FECHA Y HORA DE CREACIÓN: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROPÓSITO:
    ' Busca un valor específico dentro de un rango de celdas con coincidencia exacta
    ' y comparación case-insensitive. Optimizada para compatibilidad Excel 97-365.
    '
    ' PARÁMETROS:
    ' - rngBusqueda (Range): Rango donde realizar la búsqueda
    ' - strValorBuscado (String): Valor a buscar con coincidencia exacta
    '
    ' RETORNA: Boolean - True si encuentra el valor, False si no lo encuentra
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Variables para búsqueda
    Dim rngCelda As Range
    Dim rngEncontrado As Range
    Dim strValorCelda As String
    Dim strValorBuscadoNormalizado As String
    
    ' Inicialización
    strFuncion = "fun801_BuscarValorExactoEnRango"
    fun801_BuscarValorExactoEnRango = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validación de parámetros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    If rngBusqueda Is Nothing Then
        Exit Function
    End If
    
    If Len(Trim(strValorBuscado)) = 0 Then
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Normalización del valor buscado para comparación case-insensitive
    '--------------------------------------------------------------------------
    lngLineaError = 40
    strValorBuscadoNormalizado = UCase(Trim(strValorBuscado))
    
    '--------------------------------------------------------------------------
    ' 3. Búsqueda usando método Find (más eficiente para rangos grandes)
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Usar Find con configuración compatible Excel 97-365
    Set rngEncontrado = rngBusqueda.Find( _
        What:=strValorBuscado, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    
    ' Si Find encuentra algo, verificar que sea coincidencia exacta
    If Not rngEncontrado Is Nothing Then
        strValorCelda = UCase(Trim(CStr(rngEncontrado.Value)))
        If strValorCelda = strValorBuscadoNormalizado Then
            fun801_BuscarValorExactoEnRango = True
            Exit Function
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Método alternativo: búsqueda manual (fallback para casos especiales)
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Si Find no funcionó, usar método manual como respaldo
    For Each rngCelda In rngBusqueda.Cells
        ' Verificar que la celda no esté vacía
        If Not IsEmpty(rngCelda.Value) And Not IsNull(rngCelda.Value) Then
            strValorCelda = UCase(Trim(CStr(rngCelda.Value)))
            
            ' Comparación exacta case-insensitive
            If strValorCelda = strValorBuscadoNormalizado Then
                fun801_BuscarValorExactoEnRango = True
                Exit Function
            End If
        End If
    Next rngCelda
    
    ' Si llegamos aquí, no se encontró el valor
    fun801_BuscarValorExactoEnRango = False
    Exit Function

GestorErrores:
    ' En caso de error, retornar False
    fun801_BuscarValorExactoEnRango = False
    
    ' Log del error para debugging
    Debug.Print "Error en " & strFuncion & " línea " & lngLineaError & ": " & Err.Description
End Function

