Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_05"

Option Explicit


'******************************************************************************
' FUNCIONES AUXILIARES PARA VALIDACI�N
'******************************************************************************

Public Function fun827_ValidarNombreHoja(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun827_ValidarNombreHoja
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que el nombre de hoja sea v�lido y no est� vac�o
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun827_ValidarNombreHoja = False
    
    ' Verificar que no est� vac�o
    If Len(Trim(strNombreHoja)) = 0 Then
        Exit Function
    End If
    
    ' Verificar que no contenga caracteres no v�lidos para nombres de hoja
    If InStr(strNombreHoja, "[") > 0 Or InStr(strNombreHoja, "]") > 0 Or _
       InStr(strNombreHoja, ":") > 0 Or InStr(strNombreHoja, "*") > 0 Or _
       InStr(strNombreHoja, "?") > 0 Or InStr(strNombreHoja, "/") > 0 Or _
       InStr(strNombreHoja, "\") > 0 Then
        Exit Function
    End If
    
    ' Verificar longitud m�xima (31 caracteres para Excel)
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
    ' FUNCI�N AUXILIAR: fun828_ValidarParametrosFila
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los par�metros de fila sean v�lidos
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun828_ValidarParametrosFila = False
    
    ' Verificar que sean valores positivos
    If vFilaScenario <= 0 Or vFilaYear <= 0 Or vFilaEntity <= 0 Then
        Exit Function
    End If
    
    ' Verificar que no excedan el l�mite m�ximo de Excel (compatible con Excel 97)
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
    ' FUNCI�N AUXILIAR: fun829_ValidarParametrosColumna
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los par�metros de columna sean v�lidos
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
    
    ' Verificar que no excedan el l�mite m�ximo de Excel (compatible con Excel 97: 256 columnas)
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
    ' FUNCI�N AUXILIAR: fun830_ValidarValoresAsignar
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los valores a asignar sean v�lidos (pueden estar vac�os)
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    ' En este caso, permitimos valores vac�os ya que podr�an ser v�lidos
    ' Solo verificamos que no sean Nothing (aunque al ser String esto no aplica)
    
    ' Verificar longitud m�xima razonable para evitar problemas de memoria
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
    ' FUNCI�N AUXILIAR: fun831_ValidarRangosEnHoja
    ' Fecha y Hora de Creaci�n: 2025-06-10 03:28:21 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n: Valida que los rangos especificados existan en la hoja
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    fun831_ValidarRangosEnHoja = False
    
    ' Verificar que la hoja sea v�lida
    If ws Is Nothing Then
        Exit Function
    End If
    
    ' Verificar que las filas est�n dentro del rango de la hoja
    If vFilaScenario > ws.Rows.Count Or vFilaYear > ws.Rows.Count Or vFilaEntity > ws.Rows.Count Then
        Exit Function
    End If
    
    ' Verificar que las columnas est�n dentro del rango de la hoja
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
    ' FUNCI�N: Convertir_RangoCellsCells_a_RangoCFCF
    ' FECHA Y HORA DE CREACI�N: 2025-06-15 11:29:40 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Convierte coordenadas num�ricas de filas y columnas a formato de rango de Excel
    ' est�ndar tipo "A5:P100". Funci�n auxiliar para generaci�n din�mica de rangos
    ' de celdas en operaciones de manipulaci�n de hojas de c�lculo.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables de control de errores y validaci�n
    ' 2. Validaci�n exhaustiva de par�metros de entrada (rangos v�lidos)
    ' 3. Verificaci�n de l�gica de coordenadas (inicial <= final)
    ' 4. Conversi�n de n�meros de columna a letras usando funci�n auxiliar
    ' 5. Construcci�n del string de rango en formato Excel est�ndar
    ' 6. Validaci�n del resultado generado antes del retorno
    ' 7. Logging de operaci�n para debugging y auditor�a
    ' 8. Retorno del string de rango formateado
    ' 9. Manejo exhaustivo de errores con informaci�n detallada
    ' 10. Limpieza de recursos y logging de errores en caso de fallo
    '
    ' PAR�METROS:
    ' - vFilaInicial (Integer): N�mero de fila inicial (debe ser >= 1)
    ' - vFilaFinal (Integer): N�mero de fila final (debe ser >= vFilaInicial)
    ' - vColumnaInicial (Integer): N�mero de columna inicial (debe ser >= 1)
    ' - vColumnaFinal (Integer): N�mero de columna final (debe ser >= vColumnaInicial)
    '
    ' RETORNA: String - Rango en formato Excel (ej: "A5:P100") o cadena vac�a si error
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
    
    ' Inicializaci�n
    strFuncion = "Convertir_RangoCellsCells_a_RangoCFCF"
    Convertir_RangoCellsCells_a_RangoCFCF = ""
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables de control de errores y validaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    ' Inicializar variables de trabajo
    strColumnaInicialLetra = ""
    strColumnaFinalLetra = ""
    strRangoResultado = ""
    
    '--------------------------------------------------------------------------
    ' 2. Validaci�n exhaustiva de par�metros de entrada
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
    
    ' Validar l�mites m�ximos de Excel (compatible con Excel 97-365)
    If vFilaInicial > 65536 Then
        Err.Raise ERROR_BASE_IMPORT + 9105, strFuncion, _
            "Fila inicial excede l�mite m�ximo de Excel (65536). Valor recibido: " & vFilaInicial
    End If
    
    If vFilaFinal > 65536 Then
        Err.Raise ERROR_BASE_IMPORT + 9106, strFuncion, _
            "Fila final excede l�mite m�ximo de Excel (65536). Valor recibido: " & vFilaFinal
    End If
    
    If vColumnaInicial > 256 Then
        Err.Raise ERROR_BASE_IMPORT + 9107, strFuncion, _
            "Columna inicial excede l�mite m�ximo de Excel (256). Valor recibido: " & vColumnaInicial
    End If
    
    If vColumnaFinal > 256 Then
        Err.Raise ERROR_BASE_IMPORT + 9108, strFuncion, _
            "Columna final excede l�mite m�ximo de Excel (256). Valor recibido: " & vColumnaFinal
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Verificaci�n de l�gica de coordenadas
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
    ' 4. Conversi�n de n�meros de columna a letras
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Convertir columna inicial a letra usando funci�n auxiliar
    strColumnaInicialLetra = fun801_ConvertirNumeroColumnaALetra(vColumnaInicial)
    
    If Len(strColumnaInicialLetra) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9111, strFuncion, _
            "Error al convertir columna inicial a letra. Columna: " & vColumnaInicial
    End If
    
    ' Convertir columna final a letra usando funci�n auxiliar
    strColumnaFinalLetra = fun801_ConvertirNumeroColumnaALetra(vColumnaFinal)
    
    If Len(strColumnaFinalLetra) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9112, strFuncion, _
            "Error al convertir columna final a letra. Columna: " & vColumnaFinal
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Construcci�n del string de rango en formato Excel est�ndar
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Construir el rango en formato "COLUMNA_INICIAL+FILA_INICIAL:COLUMNA_FINAL+FILA_FINAL"
    strRangoResultado = strColumnaInicialLetra & CStr(vFilaInicial) & Chr(58) & _
                        strColumnaFinalLetra & CStr(vFilaFinal)
    
    '--------------------------------------------------------------------------
    ' 6. Validaci�n del resultado generado
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Verificar que el resultado no est� vac�o
    If Len(strRangoResultado) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9113, strFuncion, _
            "Error al generar string de rango - resultado vac�o"
    End If
    
    ' Verificar que contiene el separador de rango (:)
    If InStr(strRangoResultado, Chr(58)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 9114, strFuncion, _
            "Error en formato de rango - separador no encontrado: " & strRangoResultado
    End If
    
    ' Verificar longitud m�nima (ej: "A1:A1" = 5 caracteres m�nimo)
    If Len(strRangoResultado) < 5 Then
        Err.Raise ERROR_BASE_IMPORT + 9115, strFuncion, _
            "Longitud de rango inv�lida: " & strRangoResultado & " (Longitud: " & Len(strRangoResultado) & ")"
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Logging de operaci�n para debugging y auditor�a
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    Call fun801_LogMessage("CONVERSI�N EXITOSA - Rango generado: " & Chr(34) & strRangoResultado & Chr(34) & _
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
    ' 9. Manejo exhaustivo de errores con informaci�n detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Par�metros de entrada:" & vbCrLf & _
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
    
    ' Retornar cadena vac�a para indicar error
    Convertir_RangoCellsCells_a_RangoCFCF = ""
End Function

Public Function fun801_ConvertirNumeroColumnaALetra(ByVal vNumeroColumna As Integer) As String
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun801_ConvertirNumeroColumnaALetra
    ' FECHA Y HORA DE CREACI�N: 2025-06-15 11:29:40 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Convierte un n�mero de columna (1, 2, 3...) a su letra correspondiente
    ' en Excel (A, B, C, AA, AB...). Funci�n auxiliar para conversi�n de rangos.
    '
    ' PAR�METROS:
    ' - vNumeroColumna (Integer): N�mero de columna (1-256 para compatibilidad Excel 97)
    '
    ' RETORNA: String - Letra(s) de columna Excel o cadena vac�a si error
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim strResultado As String
    Dim intNumero As Integer
    
    ' Inicializaci�n
    fun801_ConvertirNumeroColumnaALetra = ""
    
    ' Validar par�metro
    If vNumeroColumna < 1 Or vNumeroColumna > 256 Then
        Exit Function
    End If
    
    ' Algoritmo de conversi�n a base 26 (letras A-Z)
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
    ' FUNCI�N: Contiene_Scenario_Year_Entity
    ' FECHA Y HORA DE CREACI�N: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Funci�n que verifica si una hoja espec�fica contiene tres valores exactos:
    ' escenario, a�o y sociedad. Realiza b�squeda exhaustiva en toda la hoja
    ' para determinar si los tres valores est�n presentes como contenido completo
    ' de celdas individuales.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables de control de errores y optimizaci�n
    ' 2. Validaci�n de par�metros de entrada y longitudes
    ' 3. Configuraci�n de optimizaciones de rendimiento (pantalla, c�lculo)
    ' 4. Verificaci�n de existencia de la hoja especificada
    ' 5. Obtenci�n de referencia a la hoja de trabajo
    ' 6. Determinaci�n del rango usado para b�squeda eficiente
    ' 7. B�squeda del primer valor (escenario) con coincidencia exacta
    ' 8. B�squeda del segundo valor (a�o) con coincidencia exacta
    ' 9. B�squeda del tercer valor (sociedad) con coincidencia exacta
    ' 10. Evaluaci�n de resultados y determinaci�n del valor de retorno
    ' 11. Registro de resultados en log del sistema
    ' 12. Restauraci�n de configuraciones de optimizaci�n
    ' 13. Manejo exhaustivo de errores con informaci�n detallada
    '
    ' PAR�METROS:
    ' - vSheet (String): Nombre de la hoja donde realizar la b�squeda
    ' - vEscenario (String): Valor del escenario a buscar (coincidencia exacta)
    ' - vAnio (String): Valor del a�o a buscar (coincidencia exacta)
    ' - vSociedad (String): Valor de la sociedad a buscar (coincidencia exacta)
    '
    ' VALOR DE RETORNO:
    ' - Boolean: True si los tres valores existen en la hoja, False en caso contrario
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    ' VERSI�N: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimizaci�n
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    
    ' Variables para manejo de hojas y b�squeda
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim blnHojaExiste As Boolean
    Dim rngUsedRange As Range
    
    ' Variables para resultados de b�squeda
    Dim blnExisteEscenario As Boolean
    Dim blnExisteAnio As Boolean
    Dim blnExisteSociedad As Boolean
    
    ' Variables para log de resultados
    Dim strMensajeLog As String
    
    ' Inicializaci�n
    strFuncion = "Contiene_Scenario_Year_Entity"
    Contiene_Scenario_Year_Entity = False
    lngLineaError = 0
    blnExisteEscenario = False
    blnExisteAnio = False
    blnExisteSociedad = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables de control de errores y optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando b�squeda de valores en hoja", False, "", vSheet
    
    ' Almacenar configuraciones originales para restaurar despu�s
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Validaci�n de par�metros de entrada y longitudes
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Validar que el nombre de la hoja no est� vac�o
    If Len(Trim(vSheet)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 801, strFuncion, _
            "Par�metro vSheet est� vac�o"
    End If
    
    ' Validar que el escenario no est� vac�o
    If Len(Trim(vEscenario)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 802, strFuncion, _
            "Par�metro vEscenario est� vac�o"
    End If
    
    ' Validar que el a�o no est� vac�o
    If Len(Trim(vAnio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 803, strFuncion, _
            "Par�metro vAnio est� vac�o"
    End If
    
    ' Validar que la sociedad no est� vac�a
    If Len(Trim(vSociedad)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 804, strFuncion, _
            "Par�metro vSociedad est� vac�o"
    End If
    
    ' Validar longitudes m�ximas razonables (compatibilidad Excel 97-365)
    If Len(Trim(vSheet)) > 31 Then
        Err.Raise ERROR_BASE_IMPORT + 805, strFuncion, _
            "Nombre de hoja demasiado largo: " & Len(Trim(vSheet)) & " caracteres"
    End If
    
    '--------------------------------------------------------------------------
    ' 3. Configuraci�n de optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Desactivar actualizaci�n de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar c�lculo autom�tico para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 4. Verificaci�n de existencia de la hoja especificada
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Obtener referencia al libro actual
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    ' Verificar que tenemos una referencia v�lida al libro
    If wb Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 806, strFuncion, _
            "No se pudo obtener referencia al libro de trabajo"
    End If
    
    ' Verificar existencia de la hoja usando funci�n auxiliar existente del proyecto
    blnHojaExiste = fun801_VerificarExistenciaHoja(wb, vSheet)
    
    If Not blnHojaExiste Then
        Err.Raise ERROR_BASE_IMPORT + 807, strFuncion, _
            "La hoja especificada no existe: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Obtenci�n de referencia a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 90
    Set ws = wb.Worksheets(vSheet)
    
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 808, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & vSheet
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Determinaci�n del rango usado para b�squeda eficiente
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Obtener rango usado de la hoja para optimizar b�squeda
    Set rngUsedRange = ws.UsedRange
    
    ' Verificar que la hoja tiene contenido
    If rngUsedRange Is Nothing Then
        fun801_LogMessage "Hoja est� vac�a, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    ' Verificar que el rango usado no est� vac�o
    If rngUsedRange.Cells.Count = 0 Then
        fun801_LogMessage "Rango usado est� vac�o, no hay contenido para buscar", False, "", vSheet
        GoTo RestaurarConfiguracion
    End If
    
    '--------------------------------------------------------------------------
    ' 7. B�squeda del primer valor (escenario) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 110
    
    blnExisteEscenario = fun801_BuscarValorExactoEnRango(rngUsedRange, vEscenario)
    
    fun801_LogMessage "B�squeda escenario " & Chr(34) & vEscenario & Chr(34) & _
        " resultado: " & blnExisteEscenario, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 8. B�squeda del segundo valor (a�o) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
    blnExisteAnio = fun801_BuscarValorExactoEnRango(rngUsedRange, vAnio)
    
    fun801_LogMessage "B�squeda a�o " & Chr(34) & vAnio & Chr(34) & _
        " resultado: " & blnExisteAnio, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 9. B�squeda del tercer valor (sociedad) con coincidencia exacta
    '--------------------------------------------------------------------------
    lngLineaError = 130
    
    blnExisteSociedad = fun801_BuscarValorExactoEnRango(rngUsedRange, vSociedad)
    
    fun801_LogMessage "B�squeda sociedad " & Chr(34) & vSociedad & Chr(34) & _
        " resultado: " & blnExisteSociedad, False, "", vSheet
    
    '--------------------------------------------------------------------------
    ' 10. Evaluaci�n de resultados y determinaci�n del valor de retorno
    '--------------------------------------------------------------------------
    lngLineaError = 140
    
    ' La funci�n retorna True solo si los tres valores existen
    If blnExisteEscenario And blnExisteAnio And blnExisteSociedad Then
        Contiene_Scenario_Year_Entity = True
        strMensajeLog = "�XITO - Los tres valores existen en la hoja"
    Else
        Contiene_Scenario_Year_Entity = False
        strMensajeLog = "RESULTADO - Valores faltantes: "
        If Not blnExisteEscenario Then strMensajeLog = strMensajeLog & "Escenario "
        If Not blnExisteAnio Then strMensajeLog = strMensajeLog & "A�o "
        If Not blnExisteSociedad Then strMensajeLog = strMensajeLog & "Sociedad "
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Registro de resultados en log del sistema
    '--------------------------------------------------------------------------
    lngLineaError = 150
    
    fun801_LogMessage strMensajeLog & " - Hoja: " & vSheet & _
        ", Escenario: " & Chr(34) & vEscenario & Chr(34) & _
        ", A�o: " & Chr(34) & vAnio & Chr(34) & _
        ", Sociedad: " & Chr(34) & vSociedad & Chr(34) & _
        ", Resultado final: " & Contiene_Scenario_Year_Entity, _
        False, "", vSheet

RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 12. Restauraci�n de configuraciones de optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 160
    
    ' Restaurar configuraci�n original de actualizaci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Restaurar configuraci�n original de c�lculo
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    ' Restaurar configuraci�n original de eventos
    Application.EnableEvents = blnEventsOriginal
    
    ' Limpiar referencias de objetos
    Set rngUsedRange = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    fun801_LogMessage "B�squeda completada exitosamente", False, "", vSheet
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 13. Manejo exhaustivo de errores con informaci�n detallada
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja: " & vSheet & vbCrLf & _
                      "Escenario: " & Chr(34) & vEscenario & Chr(34) & vbCrLf & _
                      "A�o: " & Chr(34) & vAnio & Chr(34) & vbCrLf & _
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
    ' FUNCI�N AUXILIAR: fun801_BuscarValorExactoEnRango
    ' FECHA Y HORA DE CREACI�N: 2025-01-16 03:00:00 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' PROP�SITO:
    ' Busca un valor espec�fico dentro de un rango de celdas con coincidencia exacta
    ' y comparaci�n case-insensitive. Optimizada para compatibilidad Excel 97-365.
    '
    ' PAR�METROS:
    ' - rngBusqueda (Range): Rango donde realizar la b�squeda
    ' - strValorBuscado (String): Valor a buscar con coincidencia exacta
    '
    ' RETORNA: Boolean - True si encuentra el valor, False si no lo encuentra
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    
    ' Variables para b�squeda
    Dim rngCelda As Range
    Dim rngEncontrado As Range
    Dim strValorCelda As String
    Dim strValorBuscadoNormalizado As String
    
    ' Inicializaci�n
    strFuncion = "fun801_BuscarValorExactoEnRango"
    fun801_BuscarValorExactoEnRango = False
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validaci�n de par�metros
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    If rngBusqueda Is Nothing Then
        Exit Function
    End If
    
    If Len(Trim(strValorBuscado)) = 0 Then
        Exit Function
    End If
    
    '--------------------------------------------------------------------------
    ' 2. Normalizaci�n del valor buscado para comparaci�n case-insensitive
    '--------------------------------------------------------------------------
    lngLineaError = 40
    strValorBuscadoNormalizado = UCase(Trim(strValorBuscado))
    
    '--------------------------------------------------------------------------
    ' 3. B�squeda usando m�todo Find (m�s eficiente para rangos grandes)
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Usar Find con configuraci�n compatible Excel 97-365
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
    ' 4. M�todo alternativo: b�squeda manual (fallback para casos especiales)
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Si Find no funcion�, usar m�todo manual como respaldo
    For Each rngCelda In rngBusqueda.Cells
        ' Verificar que la celda no est� vac�a
        If Not IsEmpty(rngCelda.Value) And Not IsNull(rngCelda.Value) Then
            strValorCelda = UCase(Trim(CStr(rngCelda.Value)))
            
            ' Comparaci�n exacta case-insensitive
            If strValorCelda = strValorBuscadoNormalizado Then
                fun801_BuscarValorExactoEnRango = True
                Exit Function
            End If
        End If
    Next rngCelda
    
    ' Si llegamos aqu�, no se encontr� el valor
    fun801_BuscarValorExactoEnRango = False
    Exit Function

GestorErrores:
    ' En caso de error, retornar False
    fun801_BuscarValorExactoEnRango = False
    
    ' Log del error para debugging
    Debug.Print "Error en " & strFuncion & " l�nea " & lngLineaError & ": " & Err.Description
End Function

