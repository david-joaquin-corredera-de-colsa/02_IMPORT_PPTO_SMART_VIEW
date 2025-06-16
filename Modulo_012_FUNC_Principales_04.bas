Attribute VB_Name = "Modulo_012_FUNC_Principales_04"
Option Explicit
'Sigue aqui: 20250609
Public Function F008_Actualizar_Informe_PL_AdHoc(ByVal strHojaPLAH As String) As Boolean
    
    '******************************************************************************
    ' Detecta donde estan los datos en la hoja del Informe PL AdHoc
    ' Modifica Scenario, Year, Entity
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Parámetros:
    ' - strInformePLAH: Nombre de la hoja del Informe de PL en formato AdHoc
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    'Variable para habilitar/deshabilitar partes de esta funcion
    Dim vEnabled_Parts As Boolean
    'vEnabled_Parts = True
    'If vEnabled_Parts Then
    'End If 'vEnabled_Parts Then
    
    ' Variables para mostrar información de rangos
    Dim strMensajeRangosDeTrabajo As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas de trabajo
    Dim wsHojaPLAH As Worksheet
    
    ' Variables para rangos de la hoja de comprobación
    Dim vFila_Inicial_HojaPLAH As Long
    Dim vFila_Final_HojaPLAH As Long
    Dim vColumna_Inicial_HojaPLAH As Long
    Dim vColumna_Final_HojaPLAH As Long
        
    
    ' Inicialización
    strFuncion = "F008_Actualizar_Informe_PL_AdHoc" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F008_Actualizar_Informe_PL_AdHoc"
    F008_Actualizar_Informe_PL_AdHoc = False
    
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para comprobar datos enviados...", False, "", strFuncion
    
    ' Validar hoja de envío
    If Not fun802_SheetExists(strHojaPLAH) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de envío no existe: " & strHojaEnvio
    End If
        
    ' Obtener referencias a las hojas
    Set wsHojaPLAH = ThisWorkbook.Worksheets(strHojaPLAH)
    
    ' Verificar que las referencias son válidas
    If wsHojaPLAH Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja del Informe PL AdHoc"
    End If
        
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave específicas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento específico
    ' Solo si necesitas valores diferentes a los por defecto
    Dim vEscenarioAdmitido, vUltimoMesCarga As String
    vEscenarioAdmitido = UCase(Trim(CONST_ESCENARIO_ADMITIDO))
    vUltimoMesCarga = UCase(Trim(CONST_ULTIMO_MES_DE_CARGA))
    Call fun826_ConfigurarPalabrasClave(vEscenarioAdmitido, vEscenarioAdmitido, vEscenarioAdmitido, vUltimoMesCarga)
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Detectando rangos de datos en hoja de Informe PL AdHoc...", False, "", strHojaComprobacion
    
    If Not fun822_DetectarRangoCompletoHoja(wsHojaPLAH, _
                                           vFila_Inicial_HojaPLAH, _
                                           vFila_Final_HojaPLAH, _
                                           vColumna_Inicial_HojaPLAH, _
                                           vColumna_Final_HojaPLAH) Then
        Err.Raise ERROR_BASE_IMPORT + 705, strFuncion, _
            "Error al detectar rangos en hoja de Informe PL AdHoc"
    End If
    
    fun801_LogMessage "Rangos de hoja Informe PL AdHoc - Filas: " & vFila_Inicial_HojaPLAH & " a " & vFila_Final_HojaPLAH & _
                      ", Columnas: " & vColumna_Inicial_HojaPLAH & " a " & vColumna_Final_HojaPLAH, _
                      False, "", strHojaPLAH
    
    vFila_Inicial_HojaPLAH = vFila_Inicial_HojaPLAH - 1 'Le quitamos 1, para que considere también la fila en la que están los headers de los meses M01 ... M12
    vColumna_Final_HojaPLAH = vColumna_Final_HojaPLAH + 22
    
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar información completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then

        lngLineaError = 125
        strMensajeRangosCompleto = "INFORMACIÓN COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE ENVÍO: " & strHojaEnvio & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaEnvio & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaEnvio & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaEnvio & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaEnvio & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaEnvio - vFila_Inicial_HojaEnvio + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaEnvio - vColumna_Inicial_HojaEnvio + 1) & vbCrLf & vbCrLf
        
        MsgBox strMensajeRangosCompleto, vbInformation, "Rangos Completos - " & strFuncion
        
    End If 'vEnabled_Parts Then
    
    
    '--------------------------------------------------------------------------
    ' 5. Procesar según el resultado de la comparación
    '--------------------------------------------------------------------------
    
    
    '----------------------------------------------------------------------
    ' 5.1. Rangos iguales: Copiar datos específicos (filas+2, columnas+11)
    '----------------------------------------------------------------------
    lngLineaError = 90003
    fun801_LogMessage "Ejecutando copia específica para rangos idénticos...", False, "", strFuncion
    
    ' Validar que hay suficientes filas y columnas para el offset
    'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
       (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
        
        ' Definir rango origen (desde comprobación)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino (hacia envío)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores únicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                "Error al copiar valores específicos"
        End If
        
        fun801_LogMessage "Copia específica completada correctamente", False, "", strFuncion
    'Else
    '    fun801_LogMessage "Advertencia: Offset insuficiente para copia específica, omitiendo operación", False, "", strFuncion
    'End If
    
        
    '--------------------------------------------------------------------------
    ' 6. Verificar integridad de la operación
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operación...", False, "", strFuncion
    
    ' Verificación básica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificación completada: ambas hojas están vacías (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificación", False, "", strFuncion
    Else
        fun801_LogMessage "Verificación completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Comprobar cada celda y etiquetar en color Verde o ROJO cada línea
    '--------------------------------------------------------------------------
    lngLineaError = 125
    fun801_LogMessage "Comprobando valores cargados en HFM" & vbCrLf & "vs valores que pretendiamos cargar en HFM ...", False, "", strFuncion
    
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    Dim vScenario As Variant
    
    'Otras variables
    Dim vColumnaEtiqueta As Integer
    vColumnaEtiqueta = 1
    Dim vValorEnviado, vValorPretendido As Double
    Dim vValorEtiqueta As String
    Dim vEtiquetaInicial, vEtiquetaOK, vEtiquetaERROR As String
    vEtiquetaInicial = "ok": vEtiquetaOK = "ok": vEtiquetaERROR = "ERROR---ERROR"
    
    
    Application.ScreenUpdating = False
    
    
    For r = vFila_Inicial_HojaComprobacion + 2 To vFila_Final_HojaComprobacion
        'Inicializamos el valor de la columna Etiqueta
        wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaInicial
        wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = xlColorIndexNone 'Sin color de fondo
        
        For c = vColumna_Inicial_HojaComprobacion + 11 To vColumna_Final_HojaComprobacion
            
            If IsNumeric(wsEnvio.Cells(r, c).Value) Then
                vValorEnviado = CDbl(wsEnvio.Cells(r, c).Value)
            Else
                vValorEnviado = wsEnvio.Cells(r, c).Value
            End If
            If IsNumeric(wsComprobacion.Cells(r, c).Value) Then
                vValorPretendido = CDbl(wsComprobacion.Cells(r, c).Value)
            Else
                vValorPretendido = wsComprobacion.Cells(r, c).Value
            End If
            vValorEtiqueta = wsEnvio.Cells(r, vColumnaEtiqueta).Value
            
            If vValorEtiqueta = vEtiquetaERROR Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaERROR
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(255, 99, 71) 'Red Tomato
            ElseIf vValorEnviado = vValorPretendido Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaOK
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(50, 205, 50) 'LimeGreen
                
            ElseIf Abs(vValorEnviado - vValorPretendido) < 0.000001 Then
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaOK
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(50, 205, 50) 'LimeGreen
            Else
                wsEnvio.Cells(r, vColumnaEtiqueta).Value = vEtiquetaERROR
                wsEnvio.Cells(r, vColumnaEtiqueta).Interior.Color = RGB(255, 99, 71) 'Red Tomato
            End If
            
        Next c
    Next r
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 6.3. Coger la Entity, FY, Scenario y llevarlo a variables especificias
    '--------------------------------------------------------------------------
    lngLineaError = 126
    fun801_LogMessage "Comprobando/almacenando valores de Entity, FY, y Scenario", False, "", strFuncion
    
    'Variables para tomar el dato/miembro del POV
    'Dim vYear_xPL As String
    'Dim vScenario_xPL As String
    'Dim vEntity_xPL As String
    
    'Variables para buscar las filas/columnas necesarias
    Dim vFilaReferencia As Integer
    Dim vColumnaReferencia As Integer
    'Variables para buscar columnas especificas
    Dim vColumnaEscenario As Integer
    Dim vColumnaYear As Integer
    Dim vColumnaEntity As Integer
    
    'Inicializamos las variables que me indican numero de fila/columna
    vFilaReferencia = vFila_Inicial_HojaComprobacion + 2
    vColumnaReferencia = vColumna_Inicial_HojaComprobacion
    vColumnaEscenario = vColumnaReferencia + 0
    vColumnaYear = vColumnaReferencia + 1
    vColumnaEntity = vColumnaReferencia + 3
    
    'Tomamos el valor para el Escenario, Year, Entity
    vScenario_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaEscenario).Value
    vYear_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaYear).Value
    vEntity_xPL = wsEnvio.Cells(vFilaReferencia, vColumnaEntity).Value
        
    '--------------------------------------------------------------------------
    ' 7. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 130
    fun801_LogMessage "Copia de datos de comprobación a envío completada con éxito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F008_Actualizar_Informe_PL_AdHoc = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuración
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F008_Actualizar_Informe_PL_AdHoc = False
End Function

Public Function F007_Preparar_Datos_para_Borrado(ByVal strHojaComprobacion As String, ByVal strHojaEnvio As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F007_Preparar_Datos_para_Borrado
    ' Fecha y Hora de Creación: 2025-06-03 00:14:44 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Copia datos específicos desde la hoja de comprobación hacia la hoja de envío,
    ' implementando lógica condicional basada en la comparación de rangos entre ambas hojas.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    ' 2. Detectar rangos de datos en hoja de comprobación
    ' 3. Detectar rangos de datos en hoja de envío
    ' 4. Comparar si los rangos son idénticos
    ' 5. Si rangos son iguales: copiar datos específicos (filas+2, columnas+11)
    ' 6. Si rangos son diferentes: copiar contenido completo y limpiar excesos
    ' 7. Verificar integridad de la operación
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Parámetros:
    ' - strHojaEnvio: Nombre de la hoja de destino (envío)
    ' - strHojaComprobacion: Nombre de la hoja de origen (comprobación)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    'Variable para habilitar/deshabilitar partes de esta funcion
    Dim vEnabled_Parts As Boolean
    
    ' Variables para mostrar información de rangos
    Dim strMensajeRangosEnvio As String
    Dim strMensajeRangosComprobacion As String
    Dim strMensajeRangosCompleto As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas de trabajo
    Dim wsEnvio As Worksheet
    Dim wsComprobacion As Worksheet
    
    ' Variables para rangos de la hoja de comprobación
    Dim vFila_Inicial_HojaComprobacion As Long
    Dim vFila_Final_HojaComprobacion As Long
    Dim vColumna_Inicial_HojaComprobacion As Long
    Dim vColumna_Final_HojaComprobacion As Long
    
    ' Variables para rangos de la hoja de envío
    Dim vFila_Inicial_HojaEnvio As Long
    Dim vFila_Final_HojaEnvio As Long
    Dim vColumna_Inicial_HojaEnvio As Long
    Dim vColumna_Final_HojaEnvio As Long
    
    ' Variable para comparación de rangos
    Dim vLosRangosSonIguales As Boolean
    
    ' Variables para rangos de copia
    Dim rngOrigen As Range
    Dim rngDestino As Range
    
    ' Inicialización
    strFuncion = "F007_Preparar_Datos_para_Borrado" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F007_Preparar_Datos_para_Borrado"
    F007_Preparar_Datos_para_Borrado = False
    lngLineaError = 0
    vLosRangosSonIguales = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar parámetros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para borrado de datos antigüos ...", False, "", strFuncion
    
    ' Validar hoja de envío
    If Not fun802_SheetExists(strHojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de envío no existe: " & strHojaEnvio
    End If
    
    ' Validar hoja de comprobación
    If Not fun802_SheetExists(strHojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 702, strFuncion, _
            "La hoja de comprobación no existe: " & strHojaComprobacion
    End If
    
    ' Obtener referencias a las hojas
    Set wsEnvio = ThisWorkbook.Worksheets(strHojaEnvio)
    Set wsComprobacion = ThisWorkbook.Worksheets(strHojaComprobacion)
    
    ' Verificar que las referencias son válidas
    If wsEnvio Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja de envío"
    End If
    
    If wsComprobacion Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 704, strFuncion, _
            "No se pudo obtener referencia a la hoja de comprobación"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave específicas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento específico
    ' Solo si necesitas valores diferentes a los por defecto
    Dim vEscenarioAdmitido, vUltimoMesCarga As String
    vEscenarioAdmitido = UCase(Trim(CONST_ESCENARIO_ADMITIDO))
    vUltimoMesCarga = UCase(Trim(CONST_ULTIMO_MES_DE_CARGA))
    Call fun826_ConfigurarPalabrasClave(vEscenarioAdmitido, vEscenarioAdmitido, vEscenarioAdmitido, vUltimoMesCarga)
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobación
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Detectando rangos de datos en hoja de comprobación...", False, "", strHojaComprobacion
    
    If Not fun822_DetectarRangoCompletoHoja(wsComprobacion, _
                                           vFila_Inicial_HojaComprobacion, _
                                           vFila_Final_HojaComprobacion, _
                                           vColumna_Inicial_HojaComprobacion, _
                                           vColumna_Final_HojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 705, strFuncion, _
            "Error al detectar rangos en hoja de comprobación"
    End If
    
    fun801_LogMessage "Rangos de comprobación - Filas: " & vFila_Inicial_HojaComprobacion & " a " & vFila_Final_HojaComprobacion & _
                      ", Columnas: " & vColumna_Inicial_HojaComprobacion & " a " & vColumna_Final_HojaComprobacion, _
                      False, "", strHojaComprobacion
    
    vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaComprobacion - 1 'Le quitamos 1, para que considere también la fila en la que están los headers de los meses M01 ... M12
    vColumna_Final_HojaComprobacion = vColumna_Inicial_HojaComprobacion + 22
    
    '--------------------------------------------------------------------------
    ' 3. Detectar rangos de datos en hoja de envío
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Detectando rangos de datos en hoja de envío...", False, "", strHojaEnvio
    
    If Not fun822_DetectarRangoCompletoHoja(wsEnvio, _
                                           vFila_Inicial_HojaEnvio, _
                                           vFila_Final_HojaEnvio, _
                                           vColumna_Inicial_HojaEnvio, _
                                           vColumna_Final_HojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 706, strFuncion, _
            "Error al detectar rangos en hoja de envío"
    End If
    
    fun801_LogMessage "Rangos de envío - Filas: " & vFila_Inicial_HojaEnvio & " a " & vFila_Final_HojaEnvio & _
                      ", Columnas: " & vColumna_Inicial_HojaEnvio & " a " & vColumna_Final_HojaEnvio, _
                      False, "", strHojaEnvio
            
    vFila_Inicial_HojaEnvio = vFila_Inicial_HojaEnvio - 1 'Le quitamos 1, para que considere también la fila en la que están los headers de los meses M01 ... M12
    vColumna_Final_HojaEnvio = vColumna_Inicial_HojaEnvio + 22
            
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar información completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then

        lngLineaError = 125
        strMensajeRangosCompleto = "INFORMACIÓN COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE ENVÍO: " & strHojaEnvio & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaEnvio & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaEnvio & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaEnvio & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaEnvio & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaEnvio - vFila_Inicial_HojaEnvio + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaEnvio - vColumna_Inicial_HojaEnvio + 1) & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE COMPROBACIÓN: " & strHojaComprobacion & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaComprobacion & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaComprobacion & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaComprobacion & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaComprobacion & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaComprobacion - vFila_Inicial_HojaComprobacion + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaComprobacion - vColumna_Inicial_HojaComprobacion + 1)
        
        MsgBox strMensajeRangosCompleto, vbInformation, "Rangos Completos - " & strFuncion
        
    End If 'vEnabled_Parts Then
    
    '--------------------------------------------------------------------------
    ' 4. Comparar si los rangos son idénticos
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Comparando rangos entre hojas...", False, "", strFuncion
    
    If (vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaEnvio) And _
       (vFila_Final_HojaComprobacion = vFila_Final_HojaEnvio) And _
       (vColumna_Inicial_HojaComprobacion = vColumna_Inicial_HojaEnvio) And _
       (vColumna_Final_HojaComprobacion = vColumna_Final_HojaEnvio) Then
        vLosRangosSonIguales = True
        fun801_LogMessage "Los rangos son idénticos - Aplicando copia específica", False, "", strFuncion
    Else
        vLosRangosSonIguales = False
        fun801_LogMessage "Los rangos son diferentes - Aplicando copia completa", False, "", strFuncion
    End If
    
    'MsgBox "Los Rangos son Iguales? = " & vLosRangosSonIguales
    
    'En realidad si los rangos no salen iguales, tiene que ser
    '   porque en una de las 2 hojas esté considerando como "Contenido"
    '   algunas celdas que en realidad no tienen contenido
    '   (tendríamos que hacerle un ClearConents a algunos rangos,
    '   como por ejemplo columnas anteriores a la del primer "BUDGET_OS", columnas posteriores a la del "M12"
    '   o filas anteriores a la del M12
    
    'Asi que vamos a forzar a que los rangos sean iguales
    ' y vamos a usar los rangos de la strHojaComprobacion
    vLosRangosSonIguales = True
    
    '--------------------------------------------------------------------------
    ' 5. Procesar según el resultado de la comparación
    '--------------------------------------------------------------------------
    If vLosRangosSonIguales = True Then
        '----------------------------------------------------------------------
        ' 5.1. Rangos iguales: Copiar datos específicos (filas+2, columnas+11)
        '----------------------------------------------------------------------
        lngLineaError = 90003
        fun801_LogMessage "Ejecutando copia específica para rangos idénticos...", False, "", strFuncion
        
        ' Validar que hay suficientes filas y columnas para el offset
        'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
           (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
            
            ' Definir rango origen (desde comprobación)
            Set rngOrigen = wsComprobacion.Range( _
                wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Definir rango destino (hacia envío)
            Set rngDestino = wsEnvio.Range( _
                wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Realizar copia de valores únicamente
            If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
                Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                    "Error al copiar valores específicos"
            End If
            
            fun801_LogMessage "Copia específica completada correctamente", False, "", strFuncion
        'Else
        '    fun801_LogMessage "Advertencia: Offset insuficiente para copia específica, omitiendo operación", False, "", strFuncion
        'End If
        
    Else
        '----------------------------------------------------------------------
        ' 5.2. Rangos diferentes: Copiar contenido completo de HojaComprobacion a HojaEnvio
        '----------------------------------------------------------------------
        lngLineaError = 100
        fun801_LogMessage "Ejecutando copia completa para rangos diferentes...", False, "", strFuncion
        
        ' Definir rango origen completo (desde comprobación)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino completo (hacia envío)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores únicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 708, strFuncion, _
                "Error al copiar contenido completo"
        End If
        
        '----------------------------------------------------------------------
        ' 5.3. Limpiar excesos en hoja de envío
        '----------------------------------------------------------------------
        lngLineaError = 110
        fun801_LogMessage "Limpiando excesos en hoja de envío...", False, "", strHojaEnvio
        
        ' Limpiar filas excedentes
        If Not fun824_LimpiarFilasExcedentes(wsEnvio, vFila_Final_HojaComprobacion) Then
            fun801_LogMessage "Advertencia: Error al limpiar filas excedentes", False, "", strHojaEnvio
        End If
        
        ' Limpiar columnas excedentes
        If Not fun825_LimpiarColumnasExcedentes(wsEnvio, vColumna_Final_HojaComprobacion) Then
            fun801_LogMessage "Advertencia: Error al limpiar columnas excedentes", False, "", strHojaEnvio
        End If
        
        fun801_LogMessage "Copia completa y limpieza completadas", False, "", strFuncion
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Verificar integridad de la operación
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operación...", False, "", strFuncion
    
    ' Verificación básica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificación completada: ambas hojas están vacías (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificación", False, "", strFuncion
    Else
        fun801_LogMessage "Verificación completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Editar cada celda para que luego el Submit pueda funcionar
    '--------------------------------------------------------------------------
    lngLineaError = 125
    fun801_LogMessage "Editando cada celda del rango para poder hacer Submit y borrar los datos antigüos...", False, "", strFuncion
    
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    Dim vScenario As Variant
    
    Application.ScreenUpdating = False
    
    For r = vFila_Inicial_HojaComprobacion + 2 To vFila_Final_HojaComprobacion
        For c = vColumna_Inicial_HojaComprobacion + 11 To vColumna_Final_HojaComprobacion
            vScenario = UCase(Trim(wsEnvio.Cells(r, vColumna_Inicial_HojaComprobacion).Value))
            If vScenario = vEscenarioAdmitido Then
                vValor = wsEnvio.Cells(r, c).Value
                wsEnvio.Cells(r, c).Value = "" 'OJO que esta linea es diferente
            Else
                'Si el escenario no es el "admitido", entonces ponemos un valor "incorrecto" para que no envíe dato
                wsEnvio.Cells(r, vColumna_Inicial_HojaComprobacion).Value = "ESCENARIO_INCORRECTO"
            End If
        Next c
    Next r
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 7. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 130
    fun801_LogMessage "Copia de datos de comprobación a envío completada con éxito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F007_Preparar_Datos_para_Borrado = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuración
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F007_Preparar_Datos_para_Borrado = False
End Function


'Public Function F009_Localizar_Hoja_Envio_Anterior(ByVal vScenario_HEnvio As String, ByVal vYear_HEnvio As String, ByVal vEntity_HEnvio As String) As Boolean
Public Function F009_Localizar_Hoja_Envio_Anterior(ByVal vScenario_HEnvio As String, ByVal vYear_HEnvio As String, ByVal vEntity_HEnvio As String) As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F009_Localizar_Hoja_Envio_Anterior
    ' Fecha y Hora de Creación: 2025-06-03 05:34:14 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Localiza la hoja de envío anterior más reciente en el libro de trabajo actual.
    ' Busca entre todas las hojas cuyo nombre comience por "Import_Envio_" y
    ' selecciona aquella con el sufijo de fecha/hora más reciente, excluyendo
    ' la hoja de envío actual.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que existe una hoja de envío actual
    ' 2. Recorrer todas las hojas del libro de trabajo
    ' 3. Identificar hojas que comienzan por "Import_Envio_"
    ' 4. Excluir la hoja de envío actual del análisis
    ' 5. Extraer y comparar sufijos de fecha/hora en formato yyyyMMdd_hhmmss
    ' 6. Seleccionar la hoja con el sufijo más reciente
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
    
    ' Inicialización
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
    ' 1. Validar que existe una hoja de envío actual
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando localización de hoja de envío anterior", False, "", strFuncion
    
    If Len(Trim(gstrNuevaHojaImportacion_Envio)) = 0 Then
        Err.Raise ERROR_BASE_IMPORT + 901, strFuncion, _
            "No se ha definido la hoja de envío actual (gstrNuevaHojaImportacion_Envio está vacía)"
    End If
    
    If Not fun802_SheetExists(gstrNuevaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 902, strFuncion, _
            "La hoja de envío actual no existe: " & gstrNuevaHojaImportacion_Envio
    End If
    
    fun801_LogMessage "Hoja de envío actual validada: " & gstrNuevaHojaImportacion_Envio, False, "", strFuncion
    
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
            ' 4. Excluir la hoja de envío actual del análisis
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
                        '20250615: aqui añadiremos la busqueda de la Entity, Scenario, Year de referencia
                        '   y las 3 tienen que ser un True para que ejecutemos las 3/4 líneas siguientes
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
        ' Declarar variable global si no existe (debería estar en el módulo de variables globales)
        gstrPreviaHojaImportacion_Envio = strHojaMayor
        
        fun801_LogMessage "Hoja de envío anterior localizada: " & gstrPreviaHojaImportacion_Envio, False, "", strFuncion
        
        '----------------------------------------------------------------------
        ' 8. Mostrar mensaje informativo
        '----------------------------------------------------------------------
        lngLineaError = 120
        MsgBox "Hoja de envío anterior localizada:" & vbCrLf & vbCrLf & _
               gstrPreviaHojaImportacion_Envio & vbCrLf & vbCrLf & _
               "Sufijo de fecha/hora: " & strSufijoMayor & vbCrLf & _
               "Esta hoja será utilizada como referencia para operaciones posteriores.", _
               vbInformation, _
               "Hoja Anterior - " & strFuncion
               
        F009_Localizar_Hoja_Envio_Anterior = True
    Else
        ' No se encontró ninguna hoja anterior
        gstrPreviaHojaImportacion_Envio = ""
        
        fun801_LogMessage "No se encontraron hojas de envío anteriores", False, "", strFuncion
        
        MsgBox "No se encontraron hojas de envío anteriores." & vbCrLf & vbCrLf & _
               "Esta parece ser la primera ejecución del proceso o " & vbCrLf & _
               "todas las hojas anteriores han sido eliminadas." & vbCrLf & vbCrLf & _
               "El proceso continuará normalmente.", _
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
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F009_Localizar_Hoja_Envio_Anterior = False
End Function

Public Function F010_Copiar_Hoja_Envio_Anterior() As Boolean
    
    '******************************************************************************
    ' FUNCIÓN PRINCIPAL: F010_Copiar_Hoja_Envio_Anterior
    ' Fecha y Hora de Creación: 2025-06-03 06:00:58 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripción:
    ' Crea una copia de la hoja de envío anterior localizada previamente
    ' y le asigna el nombre almacenado en la variable global correspondiente.
    ' Esta funcionalidad permite mantener un respaldo de la hoja anterior
    ' antes de proceder con las operaciones de importación.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar que existe una hoja de envío anterior localizada
    ' 2. Generar nombre de destino para la copia
    ' 3. Crear copia de la hoja anterior con el nuevo nombre
    ' 4. Verificar que la operación se completó correctamente
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
    
    ' Inicialización
    strFuncion = "F010_Copiar_Hoja_Envio_Anterior" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F010_Copiar_Hoja_Envio_Anterior"
    F010_Copiar_Hoja_Envio_Anterior = False
    lngLineaError = 0
    vHojaVisible = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar que existe una hoja de envío anterior localizada
    '--------------------------------------------------------------------------
    lngLineaError = 30
    fun801_LogMessage "Iniciando copia de hoja de envío anterior", False, "", strFuncion
    
    If Len(Trim(gstrPreviaHojaImportacion_Envio)) = 0 Then
        fun801_LogMessage "No hay hoja de envío anterior para copiar (primera ejecución)", False, "", strFuncion
        F010_Copiar_Hoja_Envio_Anterior = True  ' No es error, simplemente no hay hoja anterior
        Exit Function
    End If
    
    If Not fun802_SheetExists(gstrPreviaHojaImportacion_Envio) Then
        Err.Raise ERROR_BASE_IMPORT + 1001, strFuncion, _
            "La hoja de envío anterior no existe: " & gstrPreviaHojaImportacion_Envio
    End If
    
    strHojaOrigen = gstrPreviaHojaImportacion_Envio
    
    '--------------------------------------------------------------------------
    ' 2. Generar nombre de destino para la copia
    '--------------------------------------------------------------------------
    lngLineaError = 40
    If Len(Trim(gstrPrevDelHojaImportacion_Envio)) = 0 Then
        ' Generar nombre automático si no está definido
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
            "Error al copiar la hoja de envío anterior"
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Verificar que la operación se completó correctamente
    '--------------------------------------------------------------------------
    lngLineaError = 60
    If Not fun802_SheetExists(strHojaDestino) Then
        Err.Raise ERROR_BASE_IMPORT + 1003, strFuncion, _
            "Error en verificación: la hoja copiada no existe: " & strHojaDestino
    Else
        vHojaVisible = fun823_MostrarHojaSiOculta(strHojaDestino) '20250608:new
    End If
    
    '--------------------------------------------------------------------------
    ' 5. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Copia de hoja de envío anterior completada exitosamente", _
                      False, strHojaOrigen, strHojaDestino
    
    F010_Copiar_Hoja_Envio_Anterior = True
    Exit Function

GestorErrores:
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "Línea: " & lngLineaError & vbCrLf & _
                      "Número de Error: " & Err.Number & vbCrLf & _
                      "Descripción: " & Err.Description & vbCrLf & _
                      "Hoja Origen: " & strHojaOrigen & vbCrLf & _
                      "Hoja Destino: " & strHojaDestino
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F010_Copiar_Hoja_Envio_Anterior = False
End Function


