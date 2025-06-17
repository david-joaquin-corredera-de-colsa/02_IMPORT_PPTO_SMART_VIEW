Attribute VB_Name = "Modulo_012_FUNC_Principales_04"
Option Explicit
Public Function F004_Forzar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCI�N: F004_Forzar_Delimitadores_en_Excel
    ' PROP�SITO: Fuerza los delimitadores decimal y de miles en Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' PAR�METROS: Ninguno
    ' RETORNA: Boolean (True = �xito, False = error)
    '
    ' RESUMEN DE PASOS:
    ' 1. Inicializaci�n de variables globales si est�n vac�as
    ' 2. Verificaci�n de compatibilidad del sistema
    ' 3. Backup de configuraci�n actual del usuario
    ' 4. Aplicaci�n de nuevos delimitadores usando Application.International
    ' 5. Verificaci�n de aplicaci�n correcta
    ' 6. Manejo exhaustivo de errores con informaci�n detallada
    ' 7. Retorno de estado de �xito/fallo
    ' =============================================================================

    ' Variables de control de errores
    Dim strFuncionActual As String
    Dim strTipoError As String
    Dim lngLineaError As Long
    
    ' Variables de trabajo
    Dim strDelimitadorDecimalAnterior As String
    Dim strDelimitadorMilesAnterior As String
    Dim blnConfiguracionCambiada As Boolean
    
    ' Inicializaci�n
    strFuncionActual = "F004_Forzar_Delimitadores_en_Excel"
    F004_Forzar_Delimitadores_en_Excel = False
    blnConfiguracionCambiada = False
    
    On Error GoTo ErrorHandler
    
    ' =========================================================================
    ' PASO 1: Inicializaci�n de variables globales
    ' =========================================================================
    lngLineaError = 50
    Call fun801_InicializarVariablesGlobales
    
    ' =========================================================================
    ' PASO 2: Verificaci�n de compatibilidad
    ' =========================================================================
    lngLineaError = 60
    If Not fun802_VerificarCompatibilidad() Then
        strTipoError = "Error de compatibilidad del sistema"
        GoTo ErrorHandler
    End If
    
    ' =========================================================================
    ' PASO 3: Backup de configuraci�n actual
    ' =========================================================================
    lngLineaError = 70
    Call fun803_ObtenerConfiguracionActual(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
    
    ' =========================================================================
    ' PASO 4: Aplicaci�n de nuevos delimitadores
    ' =========================================================================
    lngLineaError = 80
    If fun804_AplicarNuevosDelimitadores() Then
        blnConfiguracionCambiada = True
        
        ' =====================================================================
        ' PASO 5: Verificaci�n de aplicaci�n correcta
        ' =====================================================================
        lngLineaError = 90007
        If fun805_VerificarAplicacionDelimitadores() Then
            F004_Forzar_Delimitadores_en_Excel = True
        Else
            strTipoError = "Error en verificaci�n de delimitadores aplicados"
            GoTo ErrorHandler
        End If
    Else
        strTipoError = "Error al aplicar nuevos delimitadores"
        GoTo ErrorHandler
    End If
    
    Exit Function

' =============================================================================
' CONTROL DE ERRORES EXHAUSTIVO
' =============================================================================
ErrorHandler:
    ' Restaurar configuraci�n anterior si se cambi�
    If blnConfiguracionCambiada Then
        On Error Resume Next
        Call fun806_RestaurarConfiguracion(strDelimitadorDecimalAnterior, strDelimitadorMilesAnterior)
        On Error GoTo 0
    End If
    
    ' Mostrar informaci�n detallada del error
    Call fun807_MostrarErrorDetallado(strFuncionActual, strTipoError, lngLineaError, Err.Number, Err.Description)
    
    F004_Forzar_Delimitadores_en_Excel = False
End Function




Public Function F004_Restaurar_Delimitadores_en_Excel() As Boolean

    ' =============================================================================
    ' FUNCI�N PRINCIPAL: F004_Restaurar_Delimitadores_en_Excel
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:27:06 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura los delimitadores originales de Excel desde la hoja de respaldo
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables globales con valores por defecto (C2, C3, C4)
    ' 2. Obtener referencia al libro actual
    ' 3. Verificar si existe la hoja de delimitadores originales
    ' 4. Si no existe, crear la hoja y dejarla visible (situaci�n extra�a para restauraci�n)
    ' 5. Si existe, verificar su visibilidad y hacerla visible si est� oculta
    ' 6. Leer valores originales desde las celdas especificadas:
    '    - Use System Separators desde C2
    '    - Decimal Separator desde C3
    '    - Thousands Separator desde C4
    ' 7. Almacenar valores le�dos en variables globales correspondientes
    ' 8. Validar que los valores le�dos sean apropiados para restaurar
    ' 9. Aplicar configuraci�n original de delimitadores de Excel:
    '    - Use System Separators (True/False seg�n valor original)
    '    - Decimal Separator (car�cter seg�n valor original)
    '    - Thousands Separator (car�cter seg�n valor original)
    ' 10. Verificar variable global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    ' 11. Si es True, ocultar la hoja de delimitadores al finalizar
    ' 12. Manejo exhaustivo de errores con informaci�n detallada y n�mero de l�nea
    '
    ' Par�metros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables locales
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim hojaExiste As Boolean
    Dim valorCelda As Variant
    Dim blnScreenUpdating As Boolean
    
    ' Inicializaci�n
    strFuncion = "F004_Restaurar_Delimitadores_en_Excel"
    F004_Restaurar_Delimitadores_en_Excel = False
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y CONFIGURACI�N DE ENTORNO
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando proceso de restauraci�n de delimitadores"
    fun801_LogMessage "[INICIO] " & strFuncion & " - " & strContexto, False, "", strFuncion
    fun801_LogMessage "[DETALLE] Usuario: " & Environ("USERNAME") & " | Versi�n Excel: " & Application.Version, False, "", strFuncion
    
    ' Optimizaci�n: deshabilitar actualizaciones de pantalla
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    '--------------------------------------------------------------------------
    ' PASO 2: INICIALIZAR VARIABLES GLOBALES CON VALORES POR DEFECTO
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Inicializando variables globales para delimitadores"
    fun801_LogMessage "[PASO 1] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ' Variables para las celdas que contienen los valores originales
    vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal = "C2"
    vCelda_Valor_Excel_DecimalSeparator_ValorOriginal = "C3"
    vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal = "C4"
    
    fun801_LogMessage "[DETALLE] Celdas configuradas - UseSystem: " & vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal & _
                      " | Decimal: " & vCelda_Valor_Excel_DecimalSeparator_ValorOriginal & _
                      " | Thousands: " & vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal, False, "", strFuncion
    
    ' Variables para almacenar los valores originales (inicialmente vac�as)
    vExcel_UseSystemSeparators_ValorOriginal = ""
    vExcel_DecimalSeparator_ValorOriginal = ""
    vExcel_ThousandsSeparator_ValorOriginal = ""
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: OBTENER REFERENCIA AL LIBRO ACTUAL
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Obteniendo referencia al libro de trabajo actual"
    fun801_LogMessage "[PASO 2] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
        fun801_LogMessage "[DETALLE] ThisWorkbook era Nothing, usando ActiveWorkbook", False, "", strFuncion
    End If

    If wb Is Nothing Then
        strMensajeError = "No se pudo obtener referencia v�lida al libro de trabajo"
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[�XITO] Referencia al libro obtenida - Nombre: " & wb.Name & " | Ruta: " & wb.Path, False, "", strFuncion

    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: VERIFICAR SI EXISTE LA HOJA DE DELIMITADORES ORIGINALES
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Verificando existencia de hoja de delimitadores originales"
    fun801_LogMessage "[PASO 3] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    fun801_LogMessage "[DETALLE] Buscando hoja: " & CONST_HOJA_DELIMITADORES_ORIGINALES, False, "", strFuncion
    
    hojaExiste = fun801_VerificarExistenciaHoja(wb, CONST_HOJA_DELIMITADORES_ORIGINALES)
    
    If hojaExiste Then
        fun801_LogMessage "[�XITO] Hoja de delimitadores encontrada: " & CONST_HOJA_DELIMITADORES_ORIGINALES, False, "", strFuncion
    Else
        fun801_LogMessage "[ADVERTENCIA] Hoja de delimitadores NO encontrada: " & CONST_HOJA_DELIMITADORES_ORIGINALES, False, "", strFuncion
    End If
    
    '--------------------------------------------------------------------------
    ' PASO 5: CREAR HOJA O VERIFICAR VISIBILIDAD SEG�N CORRESPONDA
    '--------------------------------------------------------------------------
    lngLineaError = 160
    
    If Not hojaExiste Then
        strContexto = "Creando hoja de delimitadores (escenario extra�o para restauraci�n)"
        fun801_LogMessage "[PASO 4A] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
        
        Set ws = fun802_CrearHojaDelimitadores(wb, CONST_HOJA_DELIMITADORES_ORIGINALES)
        If ws Is Nothing Then
            strMensajeError = "No se pudo crear la hoja de delimitadores originales: " & CONST_HOJA_DELIMITADORES_ORIGINALES
            Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, strMensajeError
        End If
        
        fun801_LogMessage "[ADVERTENCIA] Hoja creada pero no hay valores para restaurar - proceso finalizado exitosamente", False, "", strFuncion
        F004_Restaurar_Delimitadores_en_Excel = True
        Application.ScreenUpdating = blnScreenUpdating
        Exit Function
    Else
        lngLineaError = 170
        strContexto = "Obteniendo referencia a hoja existente y verificando visibilidad"
        fun801_LogMessage "[PASO 4B] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
        
        Set ws = wb.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES)
        fun801_LogMessage "[DETALLE] Referencia a hoja obtenida - Estado visible actual: " & ws.Visible, False, "", strFuncion
        
        ' Verificar si est� oculta y hacerla visible si es necesario
        If Not fun803_HacerHojaVisible(ws) Then
            fun801_LogMessage "[ADVERTENCIA] No se pudo hacer visible la hoja " & CONST_HOJA_DELIMITADORES_ORIGINALES & _
                              " (L�nea: " & lngLineaError & ")", False, "", strFuncion
        Else
            fun801_LogMessage "[�XITO] Hoja configurada como visible: " & CONST_HOJA_DELIMITADORES_ORIGINALES, False, "", strFuncion
        End If
    End If
    
    lngLineaError = 180
    
    '--------------------------------------------------------------------------
    ' PASO 6: LEER VALORES ORIGINALES DESDE LAS CELDAS ESPECIFICADAS
    '--------------------------------------------------------------------------
    lngLineaError = 190
    strContexto = "Leyendo valores originales desde celdas especificadas"
    fun801_LogMessage "[PASO 5] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ' Leer valor de Use System Separators desde C2
    lngLineaError = 200
    On Error Resume Next
    valorCelda = ws.Range(vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal).Value
    If Err.Number <> 0 Then
        fun801_LogMessage "[ERROR] Error al leer celda " & vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal & _
                          " - Error: " & Err.Number & " - " & Err.Description & " (L�nea: " & lngLineaError & ")", True, "", strFuncion
        On Error GoTo GestorErrores
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, "Error al leer UseSystemSeparators desde " & vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal
    End If
    On Error GoTo GestorErrores
    
    vExcel_UseSystemSeparators_ValorOriginal = fun804_ConvertirValorACadena(valorCelda)
    fun801_LogMessage "[DETALLE] UseSystemSeparators le�do - Celda: " & vCelda_Valor_Excel_UseSystemSeparators_ValorOriginal & _
                      " | Valor raw: " & CStr(valorCelda) & " | Valor convertido: " & vExcel_UseSystemSeparators_ValorOriginal, False, "", strFuncion
    
    ' Leer valor de Decimal Separator desde C3
    lngLineaError = 210
    On Error Resume Next
    valorCelda = ws.Range(vCelda_Valor_Excel_DecimalSeparator_ValorOriginal).Value
    If Err.Number <> 0 Then
        fun801_LogMessage "[ERROR] Error al leer celda " & vCelda_Valor_Excel_DecimalSeparator_ValorOriginal & _
                          " - Error: " & Err.Number & " - " & Err.Description & " (L�nea: " & lngLineaError & ")", True, "", strFuncion
        On Error GoTo GestorErrores
        Err.Raise ERROR_BASE_IMPORT + 1004, strFuncion, "Error al leer DecimalSeparator desde " & vCelda_Valor_Excel_DecimalSeparator_ValorOriginal
    End If
    On Error GoTo GestorErrores
    
    vExcel_DecimalSeparator_ValorOriginal = fun804_ConvertirValorACadena(valorCelda)
    fun801_LogMessage "[DETALLE] DecimalSeparator le�do - Celda: " & vCelda_Valor_Excel_DecimalSeparator_ValorOriginal & _
                      " | Valor raw: " & CStr(valorCelda) & " | Valor convertido: " & vExcel_DecimalSeparator_ValorOriginal, False, "", strFuncion
    
    ' Leer valor de Thousands Separator desde C4
    lngLineaError = 220
    On Error Resume Next
    valorCelda = ws.Range(vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal).Value
    If Err.Number <> 0 Then
        fun801_LogMessage "[ERROR] Error al leer celda " & vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal & _
                          " - Error: " & Err.Number & " - " & Err.Description & " (L�nea: " & lngLineaError & ")", True, "", strFuncion
        On Error GoTo GestorErrores
        Err.Raise ERROR_BASE_IMPORT + 1005, strFuncion, "Error al leer ThousandsSeparator desde " & vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal
    End If
    On Error GoTo GestorErrores
    
    vExcel_ThousandsSeparator_ValorOriginal = fun804_ConvertirValorACadena(valorCelda)
    fun801_LogMessage "[DETALLE] ThousandsSeparator le�do - Celda: " & vCelda_Valor_Excel_ThousandsSeparator_ValorOriginal & _
                      " | Valor raw: " & CStr(valorCelda) & " | Valor convertido: " & vExcel_ThousandsSeparator_ValorOriginal, False, "", strFuncion

    lngLineaError = 230
    
    '--------------------------------------------------------------------------
    ' PASO 7: VALIDAR QUE SE HAYAN LE�DO VALORES V�LIDOS
    '--------------------------------------------------------------------------
    lngLineaError = 240
    strContexto = "Validando valores le�dos para restauraci�n"
    fun801_LogMessage "[PASO 6] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    fun801_LogMessage "[DETALLE] Valores a validar - UseSystem: '" & vExcel_UseSystemSeparators_ValorOriginal & _
                      "' | Decimal: '" & vExcel_DecimalSeparator_ValorOriginal & _
                      "' | Thousands: '" & vExcel_ThousandsSeparator_ValorOriginal & "'", False, "", strFuncion
    
    If Not fun805_ValidarValoresOriginales() Then
        strMensajeError = "No se encontraron valores v�lidos para restaurar en la hoja: " & CONST_HOJA_DELIMITADORES_ORIGINALES & _
                         " | UseSystem: '" & vExcel_UseSystemSeparators_ValorOriginal & _
                         "' | Decimal: '" & vExcel_DecimalSeparator_ValorOriginal & _
                         "' | Thousands: '" & vExcel_ThousandsSeparator_ValorOriginal & "'"
        Err.Raise ERROR_BASE_IMPORT + 1006, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[�XITO] Validaci�n de valores completada exitosamente", False, "", strFuncion
    
    lngLineaError = 250
    
    '--------------------------------------------------------------------------
    ' PASO 8: APLICAR CONFIGURACI�N ORIGINAL DE DELIMITADORES DE EXCEL
    '--------------------------------------------------------------------------
    lngLineaError = 260
    strContexto = "Aplicando configuraci�n original de delimitadores de Excel"
    fun801_LogMessage "[PASO 7] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ' Restaurar Use System Separators (True/False)
    lngLineaError = 270
    fun801_LogMessage "[SUB-PASO 7A] Restaurando UseSystemSeparators: '" & vExcel_UseSystemSeparators_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    If Not fun806_RestaurarUseSystemSeparators(vExcel_UseSystemSeparators_ValorOriginal) Then
        fun801_LogMessage "[ADVERTENCIA] Error al restaurar UseSystemSeparators - Valor: '" & vExcel_UseSystemSeparators_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    Else
        fun801_LogMessage "[�XITO] UseSystemSeparators restaurado exitosamente", False, "", strFuncion
    End If
    
    ' Restaurar Decimal Separator (car�cter)
    lngLineaError = 280
    fun801_LogMessage "[SUB-PASO 7B] Restaurando DecimalSeparator: '" & vExcel_DecimalSeparator_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    If Not fun807_RestaurarDecimalSeparator(vExcel_DecimalSeparator_ValorOriginal) Then
        fun801_LogMessage "[ADVERTENCIA] Error al restaurar DecimalSeparator - Valor: '" & vExcel_DecimalSeparator_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    Else
        fun801_LogMessage "[�XITO] DecimalSeparator restaurado exitosamente", False, "", strFuncion
    End If
    
    ' Restaurar Thousands Separator (car�cter)
    lngLineaError = 290
    fun801_LogMessage "[SUB-PASO 7C] Restaurando ThousandsSeparator: '" & vExcel_ThousandsSeparator_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    If Not fun808_RestaurarThousandsSeparator(vExcel_ThousandsSeparator_ValorOriginal) Then
        fun801_LogMessage "[ADVERTENCIA] Error al restaurar ThousandsSeparator - Valor: '" & vExcel_ThousandsSeparator_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    Else
        fun801_LogMessage "[�XITO] ThousandsSeparator restaurado exitosamente", False, "", strFuncion
    End If
    
    lngLineaError = 300
    
    '--------------------------------------------------------------------------
    ' PASO 9: VERIFICAR SI DEBE OCULTAR LA HOJA
    '--------------------------------------------------------------------------
    lngLineaError = 310
    strContexto = "Configurando visibilidad final de la hoja de delimitadores"
    fun801_LogMessage "[PASO 8] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    fun801_LogMessage "[DETALLE] Configuraci�n de visibilidad: CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE = " & CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE, False, "", strFuncion
    
    ' Configurar visibilidad seg�n constante global
    ThisWorkbook.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES).Visible = CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE
    fun801_LogMessage "[�XITO] Visibilidad de hoja configurada seg�n constante global", False, "", strFuncion
    
    lngLineaError = 320
    
    '--------------------------------------------------------------------------
    ' PASO 10: FINALIZACI�N EXITOSA
    '--------------------------------------------------------------------------
    lngLineaError = 330
    strContexto = "Finalizando proceso de restauraci�n exitosamente"
    fun801_LogMessage "[PASO 9] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ' Verificar delimitadores aplicados actualmente
    fun801_LogMessage "[VERIFICACI�N FINAL] Delimitadores actuales - Decimal: '" & Application.DecimalSeparator & _
                      "' | Thousands: '" & Application.ThousandsSeparator & "'", False, "", strFuncion
    
    ' Restaurar configuraci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdating
    
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Total l�neas procesadas: " & lngLineaError, False, "", strFuncion
    F004_Restaurar_Delimitadores_en_Excel = True
    
    Exit Function
    
GestorErrores:
    ' Restaurar configuraci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdating
    
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Fuente del Error: " & Err.Source & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Versi�n Excel: " & Application.Version & vbCrLf & _
                      "Libro de Trabajo: " & IIf(wb Is Nothing, "Nothing", wb.Name) & vbCrLf & _
                      "Hoja de Delimitadores: " & CONST_HOJA_DELIMITADORES_ORIGINALES & vbCrLf & _
                      "Estados de Variables Globales:" & vbCrLf & _
                      "  - UseSystemSeparators: '" & vExcel_UseSystemSeparators_ValorOriginal & "'" & vbCrLf & _
                      "  - DecimalSeparator: '" & vExcel_DecimalSeparator_ValorOriginal & "'" & vbCrLf & _
                      "  - ThousandsSeparator: '" & vExcel_ThousandsSeparator_ValorOriginal & "'" & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    F004_Restaurar_Delimitadores_en_Excel = False
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
End Function


Public Function F009_Localizar_Hoja_Envio_Anterior(ByVal vScenario_HEnvio As String, ByVal vYear_HEnvio As String, ByVal vEntity_HEnvio As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N PRINCIPAL: F009_Localizar_Hoja_Envio_Anterior
    ' Fecha y Hora de Creaci�n: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Localiza la hoja de env�o anterior m�s reciente en el libro de trabajo actual.
    ' Busca entre todas las hojas cuyo nombre comience por "Import_Envio_" y
    ' selecciona aquella con el sufijo de fecha/hora m�s reciente, excluyendo
    ' la hoja de env�o actual.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que existe una hoja de env�o actual
    ' 2. Recorrer todas las hojas del libro de trabajo
    ' 3. Identificar hojas que comienzan por "Import_Envio_"
    ' 4. Excluir la hoja de env�o actual del an�lisis
    ' 5. Extraer y comparar sufijos de fecha/hora en formato yyyyMMdd_hhmmss
    ' 6. Seleccionar la hoja con el sufijo m�s reciente
    ' 7. Almacenar el resultado en variable global gstrPreviaHojaImportacion_Envio
    ' 8. Mostrar mensaje informativo con la hoja seleccionada
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para procesamiento
    Dim ws As Worksheet
    Dim strNombreHoja As String
    
    Dim strSufijoActual As String
    Dim strSufijoMayor As String
    Dim strHojaMayor As String
    Dim intLongitudSufijo As Integer
    Dim blnEncontradaHoja As Boolean
    
    ' Inicializaci�n
    strFuncion = "F009_Localizar_Hoja_Envio_Anterior" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F009_Localizar_Hoja_Envio_Anterior"
    F009_Localizar_Hoja_Envio_Anterior = False
    lngLineaError = 0
    
    ' Constantes de trabajo
    
    intLongitudSufijo = 15  ' yyyyMMdd_hhmmss = 15 caracteres
    strSufijoMayor = ""
    strHojaMayor = ""
    blnEncontradaHoja = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que existe una hoja de env�o actual
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando localizaci�n de hoja de env�o anterior", False, "", strFuncion
    
    If Len(Trim(gstrNuevaHojaImportacion_Envio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 901, strFuncion, _
            "No se ha definido la hoja de env�o actual (gstrNuevaHojaImportacion_Envio est� vac�a)"
    End If
    
    If Not fun802_SheetExists(gstrNuevaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 902, strFuncion, _
            "La hoja de env�o actual no existe: " & gstrNuevaHojaImportacion_Envio
    End If
    
    fun801_LogMessage "Hoja de env�o actual validada: " & gstrNuevaHojaImportacion_Envio, False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 2. Recorrer todas las hojas del libro de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Iniciando recorrido de hojas del libro", False, "", strFuncion
        
    For Each ws In ThisWorkbook.Worksheets
        strNombreHoja = ws.Name
        
        '----------------------------------------------------------------------
        ' 3. Identificar hojas que comienzan por "Import_Envio_"
        '----------------------------------------------------------------------
        lngLineaError = 70
        If fun821_ComenzarPorPrefijo(strNombreHoja, CONST_PREFIJO_HOJA_IMPORTACION_ENVIO) Then
            
            '------------------------------------------------------------------
            ' 4. Excluir la hoja de env�o actual del an�lisis
            '------------------------------------------------------------------
            lngLineaError = 80
            If strNombreHoja <> gstrNuevaHojaImportacion_Envio Then
                
                '--------------------------------------------------------------
                ' 5. Extraer y validar sufijo de fecha/hora (para hojas que comienzan por "Import_Envio_")
                '--------------------------------------------------------------
                lngLineaError = 90005
                If fun822_ValidarFormatoSufijoFecha(strNombreHoja, CONST_PREFIJO_HOJA_IMPORTACION_ENVIO, intLongitudSufijo) Then
                    
                    ' Extraer sufijo
                    strSufijoActual = fun823_ExtraerSufijoFecha(strNombreHoja, intLongitudSufijo)
                    
                    '----------------------------------------------------------
                    ' 6. Comparar sufijos y seleccionar el mayor
                    '----------------------------------------------------------
                    lngLineaError = 100
                    If fun824_CompararSufijosFecha(strSufijoActual, strSufijoMayor) > 0 Then
                        '20250615: aqui a�adiremos la busqueda de la Entity, Scenario, Year de referencia
                        '   y las 3 tienen que ser un True para que ejecutemos las 3/4 l�neas siguientes
                        If Contiene_Scenario_Year_Entity(strNombreHoja, vScenario_HEnvio, vYear_HEnvio, vEntity_HEnvio) Then '20250615
                            strSufijoMayor = strSufijoActual
                            strHojaMayor = strNombreHoja
                            blnEncontradaHoja = True
                            
                            fun801_LogMessage "Nueva hoja candidata encontrada: " & strNombreHoja & " (Sufijo: " & strSufijoActual & ")", _
                                              False, "", strFuncion
                        
                        End If '20250615
                    End If
                End If
            Else
                fun801_LogMessage "Hoja excluida (es la actual): " & strNombreHoja, False, "", strFuncion
            End If
        End If
    Next ws
    
    '--------------------------------------------------------------------------
    ' 7. Almacenar resultado en variable global
    '--------------------------------------------------------------------------
    lngLineaError = 110
    If blnEncontradaHoja Then
        ' Declarar variable global si no existe (deber�a estar en el m�dulo de variables globales)
        gstrPreviaHojaImportacion_Envio = strHojaMayor
        
        fun801_LogMessage "Hoja de env�o anterior localizada: " & gstrPreviaHojaImportacion_Envio, False, "", strFuncion
        
        '----------------------------------------------------------------------
        ' 8. Mostrar mensaje informativo
        '----------------------------------------------------------------------
        lngLineaError = 120
        MsgBox "Hoja de env�o anterior localizada:" & vbCrLf & vbCrLf & _
               gstrPreviaHojaImportacion_Envio & vbCrLf & vbCrLf & _
               "Sufijo de fecha/hora: " & strSufijoMayor & vbCrLf & _
               "Esta hoja ser� utilizada como referencia para operaciones posteriores.", _
               vbInformation, _
               "Hoja Anterior - " & strFuncion
               
        F009_Localizar_Hoja_Envio_Anterior = True
    Else
        ' No se encontr� ninguna hoja anterior
        gstrPreviaHojaImportacion_Envio = ""
        
        fun801_LogMessage "No se encontraron hojas de env�o anteriores", False, "", strFuncion
        
        MsgBox "No se encontraron hojas de env�o anteriores." & vbCrLf & vbCrLf & _
               "Esta parece ser la primera ejecuci�n del proceso o " & vbCrLf & _
               "todas las hojas anteriores han sido eliminadas." & vbCrLf & vbCrLf & _
               "El proceso continuar� normalmente.", _
               vbInformation, _
               "Sin Hojas Anteriores - " & strFuncion
               
        F009_Localizar_Hoja_Envio_Anterior = True ' No es error, simplemente no hay hojas anteriores
        'Si no hemos encontrado hoja previa usamos la nueva hoja importacion envio en su lugar
        
        gstrPreviaHojaImportacion_Envio = gstrNuevaHojaImportacion_Envio
    End If
    
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F009_Localizar_Hoja_Envio_Anterior = False
End Function

Public Function F010_Copiar_Hoja_Envio_Anterior() As Boolean
    
    '******************************************************************************
    ' FUNCI�N PRINCIPAL: F010_Copiar_Hoja_Envio_Anterior
    ' Fecha y Hora de Creaci�n: 2025-06-03 06:00:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Crea una copia de la hoja de env�o anterior localizada previamente
    ' y le asigna el nombre almacenado en la variable global correspondiente.
    ' Esta funcionalidad permite mantener un respaldo de la hoja anterior
    ' antes de proceder con las operaciones de importaci�n.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que existe una hoja de env�o anterior localizada
    ' 2. Generar nombre de destino para la copia
    ' 3. Crear copia de la hoja anterior con el nuevo nombre
    ' 4. Verificar que la operaci�n se complet� correctamente
    ' 5. Registrar resultado en el log del sistema
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    'Variable para mostrar hoja si oculta
    Dim vHojaVisible As Boolean
    
    ' Variables para procesamiento
    Dim strHojaOrigen As String
    Dim strHojaDestino As String
    
    ' Inicializaci�n
    strFuncion = "F010_Copiar_Hoja_Envio_Anterior" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F010_Copiar_Hoja_Envio_Anterior"
    F010_Copiar_Hoja_Envio_Anterior = False
    lngLineaError = 0
    vHojaVisible = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que existe una hoja de env�o anterior localizada
    '--------------------------------------------------------------------------
    lngLineaError = 30
    fun801_LogMessage "Iniciando copia de hoja de env�o anterior", False, "", strFuncion
    
    If Len(Trim(gstrPreviaHojaImportacion_Envio)) = 0 Then
        fun801_LogMessage "No hay hoja de env�o anterior para copiar (primera ejecuci�n)", False, "", strFuncion
        F010_Copiar_Hoja_Envio_Anterior = True  ' No es error, simplemente no hay hoja anterior
        Exit Function
    End If
    
    If Not fun802_SheetExists(gstrPreviaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "La hoja de env�o anterior no existe: " & gstrPreviaHojaImportacion_Envio
    End If
    
    strHojaOrigen = gstrPreviaHojaImportacion_Envio
    
    '--------------------------------------------------------------------------
    ' 2. Generar nombre de destino para la copia
    '--------------------------------------------------------------------------
    lngLineaError = 40
    If Len(Trim(gstrPrevDelHojaImportacion_Envio)) = 0 Then
        ' Generar nombre autom�tico si no est� definido
        gstrPrevDelHojaImportacion_Envio = CONST_PREFIJO_BACKUP_HOJA_PREVIA_ENVIO & strHojaOrigen
    End If
    
    strHojaDestino = gstrPrevDelHojaImportacion_Envio
    
    fun801_LogMessage "Preparando copia: " & strHojaOrigen & " ? " & strHojaDestino, False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 3. Crear copia de la hoja anterior
    '--------------------------------------------------------------------------
    lngLineaError = 50
    If Not fun825_CopiarHojaConNuevoNombre(strHojaOrigen, strHojaDestino) Then
        Err.Raise ERROR_BASE_IMPORT + 1002, strFuncion, _
            "Error al copiar la hoja de env�o anterior"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Verificar que la operaci�n se complet� correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 60
    If Not fun802_SheetExists(strHojaDestino) Then
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, _
            "Error en verificaci�n: la hoja copiada no existe: " & strHojaDestino
    Else
        vHojaVisible = fun823_MostrarHojaSiOculta(strHojaDestino) '20250608:new
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Copia de hoja de env�o anterior completada exitosamente", _
                      False, strHojaOrigen, strHojaDestino
    
    F010_Copiar_Hoja_Envio_Anterior = True
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja Origen: " & strHojaOrigen & vbCrLf & _
                      "Hoja Destino: " & strHojaDestino
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F010_Copiar_Hoja_Envio_Anterior = False
End Function
Public Function F004_Detectar_Delimitadores_en_Excel() As Boolean
    
    ' =============================================================================
    ' FUNCI�N PRINCIPAL: F004_Detectar_Delimitadores_en_Excel
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-05-26 17:43:59 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    ' Descripci�n: Detecta y almacena los delimitadores de Excel actuales
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializar variables globales con valores por defecto
    ' 2. Verificar si existe la hoja de delimitadores originales
    ' 3. Si no existe, crear la hoja y dejarla visible
    ' 4. Si existe, verificar su visibilidad y hacerla visible si est� oculta
    ' 5. Limpiar el contenido de la hoja una vez visible
    ' 6. Configurar headers en las celdas especificadas (B2, B3, B4)
    ' 7. Detectar configuraci�n actual de delimitadores de Excel:
    '    - Use System Separators (True/False)
    '    - Decimal Separator (car�cter)
    '    - Thousands Separator (car�cter)
    ' 8. Almacenar valores detectados en variables globales
    ' 9. Escribir valores en la hoja de delimitadores (C2, C3, C4)
    ' 10. Verificar constante global CONST_OCULTAR_REPOSITORIO_DELIMITADORES
    ' 11. Si es True, ocultar la hoja creada/actualizada
    ' 12. Manejo exhaustivo de errores con informaci�n detallada
    '
    ' Par�metros: Ninguno
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    ' Control de errores con n�mero de l�nea
    On Error GoTo ErrorHandler
    
    ' Variables locales
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim hojaExiste As Boolean
    Dim i As Integer
    Dim lineaError As Long
    
    ' Inicializar resultado como exitoso
    F004_Detectar_Delimitadores_en_Excel = True
    
    ' ==========================================================================
    ' PASO 1: INICIALIZAR VARIABLES GLOBALES CON VALORES POR DEFECTO
    ' ==========================================================================
    lineaError = 100
        
    ' Celdas para los headers (t�tulos)
    vCelda_Header_Excel_UseSystemSeparators = "B2"
    vCelda_Header_Excel_DecimalSeparator = "B3"
    vCelda_Header_Excel_ThousandsSeparator = "B4"
    
    ' Celdas para los valores detectados
    vCelda_Valor_Excel_UseSystemSeparators = "C2"
    vCelda_Valor_Excel_DecimalSeparator = "C3"
    vCelda_Valor_Excel_ThousandsSeparator = "C4"
    
    ' Variables para almacenar los valores detectados (inicialmente vac�as)
    vExcel_UseSystemSeparators = ""
    vExcel_DecimalSeparator = ""
    vExcel_ThousandsSeparator = ""
    
    lineaError = 110
    
    ' ==========================================================================
    ' PASO 2: OBTENER REFERENCIA AL LIBRO ACTUAL
    ' ==========================================================================
    
'    Set wb = ActiveWorkbook
'    If wb Is Nothing Then
'        Set wb = ThisWorkbook
'    End If
'
'    If wb Is Nothing Then
'        F004_Detectar_Delimitadores_en_Excel = False
'        Exit Function
'    End If
'
    lineaError = 120
    
    ' ==========================================================================
    ' PASO 3: VERIFICAR SI EXISTE LA HOJA DE DELIMITADORES ORIGINALES
    ' ==========================================================================
    
'    hojaExiste = fun801_VerificarExistenciaHoja(wb, CONST_HOJA_DELIMITADORES_ORIGINALES)
    
    lineaError = 130
    
    ' ==========================================================================
    ' PASO 4: CREAR HOJA O VERIFICAR VISIBILIDAD SEG�N CORRESPONDA
    ' ==========================================================================
    
'    If Not hojaExiste Then
'        ' La hoja no existe, crearla y dejarla visible
'        Set ws = fun802_CrearHojaDelimitadores(wb, CONST_HOJA_DELIMITADORES_ORIGINALES)
'        If ws Is Nothing Then
'            F004_Detectar_Delimitadores_en_Excel = False
'            Exit Function
'        End If
'        ' La hoja reci�n creada ya est� visible por defecto
'    Else
        ' La hoja existe, obtener referencia y verificar visibilidad
'        Set ws = wb.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES)
        Set wb = ThisWorkbook
        Set ws = ThisWorkbook.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES)
        ThisWorkbook.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES).Visible = xlSheetVisible
        
        ' Verificar si est� oculta y hacerla visible si es necesario
'        Call fun803_HacerHojaVisible(ws)
        
'    End If
    
    lineaError = 140
    
    ' ==========================================================================
    ' PASO 5: LIMPIAR CONTENIDO DE LA HOJA (AHORA QUE EST� VISIBLE)
    ' ==========================================================================
    
    Call fun804_LimpiarContenidoHoja(ws)
    
    lineaError = 150
    
    ' ==========================================================================
    ' PASO 6: CONFIGURAR HEADERS EN LAS CELDAS ESPECIFICADAS
    ' ==========================================================================
    
    ' Header para Use System Separators en B2
    ws.Range(vCelda_Header_Excel_UseSystemSeparators).Value = "Excel Use System Separators"
    
    ' Header para Decimal Separator en B3
    ws.Range(vCelda_Header_Excel_DecimalSeparator).Value = "Excel Decimals"
    
    ' Header para Thousands Separator en B4
    ws.Range(vCelda_Header_Excel_ThousandsSeparator).Value = "Excel Thousands"
    
    lineaError = 160
    
    ' ==========================================================================
    ' PASO 7: DETECTAR CONFIGURACI�N ACTUAL DE DELIMITADORES DE EXCEL
    ' ==========================================================================
    
    ' Detectar Use System Separators
    vExcel_UseSystemSeparators = fun805_DetectarUseSystemSeparators()
    
    ' Detectar Decimal Separator
    vExcel_DecimalSeparator = fun806_DetectarDecimalSeparator()
    
    ' Detectar Thousands Separator
    vExcel_ThousandsSeparator = fun807_DetectarThousandsSeparator()
    
    lineaError = 170
    
    ' ==========================================================================
    ' PASO 8: ALMACENAR VALORES DETECTADOS EN LA HOJA
    ' ==========================================================================
    
    ' Almacenar Use System Separators en C2
    ws.Range(vCelda_Valor_Excel_UseSystemSeparators).Value = vExcel_UseSystemSeparators
    
    ' Almacenar Decimal Separator en C3
    ws.Range(vCelda_Valor_Excel_DecimalSeparator).Value = vExcel_DecimalSeparator
    
    ' Almacenar Thousands Separator en C4
    ws.Range(vCelda_Valor_Excel_ThousandsSeparator).Value = vExcel_ThousandsSeparator
    
    lineaError = 180
    
    ' ==========================================================================
    ' PASO 9: VERIFICAR SI DEBE OCULTAR LA HOJA
    ' ==========================================================================
    
    ThisWorkbook.Worksheets(CONST_HOJA_DELIMITADORES_ORIGINALES).Visible = CONST_HOJA_DELIMITADORES_ORIGINALES_VISIBLE
    lineaError = 190
    
    ' ==========================================================================
    ' PASO 10: FINALIZACI�N EXITOSA
    ' ==========================================================================
    
    Exit Function
    
ErrorHandler:
    ' ==========================================================================
    ' MANEJO EXHAUSTIVO DE ERRORES
    ' ==========================================================================
    
    F004_Detectar_Delimitadores_en_Excel = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: F004_Detectar_Delimitadores_en_Excel" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now() & vbCrLf & _
                   "USUARIO: david-joaquin-corredera-de-colsa"
    
    ' Mostrar mensaje de error (comentar si no se desea)
    ' MsgBox mensajeError, vbCritical, "Error en Detecci�n de Delimitadores"
    
    ' Log del error para debugging
    Debug.Print mensajeError
    
    ' Limpiar objetos
    Set ws = Nothing
    Set wb = Nothing
    
End Function


