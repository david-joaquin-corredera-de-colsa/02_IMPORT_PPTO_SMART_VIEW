Attribute VB_Name = "Modulo_012_FUNC_Principales_04"
Option Explicit
'Sigue aqui: 20250609
Public Function F008_Actualizar_Informe_PL_AdHoc(ByVal strHojaPLAH As String) As Boolean
    
    '******************************************************************************
    ' Detecta donde estan los datos en la hoja del Informe PL AdHoc
    ' Modifica Scenario, Year, Entity
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Par�metros:
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
    
    ' Variables para mostrar informaci�n de rangos
    Dim strMensajeRangosDeTrabajo As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para hojas de trabajo
    Dim wsHojaPLAH As Worksheet
    
    ' Variables para rangos de la hoja de comprobaci�n
    Dim vFila_Inicial_HojaPLAH As Long
    Dim vFila_Final_HojaPLAH As Long
    Dim vColumna_Inicial_HojaPLAH As Long
    Dim vColumna_Final_HojaPLAH As Long
        
    
    ' Inicializaci�n
    strFuncion = "F008_Actualizar_Informe_PL_AdHoc" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F008_Actualizar_Informe_PL_AdHoc"
    F008_Actualizar_Informe_PL_AdHoc = False
    
    lngLineaError = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para comprobar datos enviados...", False, "", strFuncion
    
    ' Validar hoja de env�o
    If Not fun802_SheetExists(strHojaPLAH) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de env�o no existe: " & strHojaEnvio
    End If
        
    ' Obtener referencias a las hojas
    Set wsHojaPLAH = ThisWorkbook.Worksheets(strHojaPLAH)
    
    ' Verificar que las referencias son v�lidas
    If wsHojaPLAH Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja del Informe PL AdHoc"
    End If
        
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave espec�ficas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento espec�fico
    ' Solo si necesitas valores diferentes a los por defecto
    Dim vEscenarioAdmitido, vUltimoMesCarga As String
    vEscenarioAdmitido = UCase(Trim(CONST_ESCENARIO_ADMITIDO))
    vUltimoMesCarga = UCase(Trim(CONST_ULTIMO_MES_DE_CARGA))
    Call fun826_ConfigurarPalabrasClave(vEscenarioAdmitido, vEscenarioAdmitido, vEscenarioAdmitido, vUltimoMesCarga)
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobaci�n
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
    
    vFila_Inicial_HojaPLAH = vFila_Inicial_HojaPLAH - 1 'Le quitamos 1, para que considere tambi�n la fila en la que est�n los headers de los meses M01 ... M12
    vColumna_Final_HojaPLAH = vColumna_Final_HojaPLAH + 22
    
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar informaci�n completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then

        lngLineaError = 125
        strMensajeRangosCompleto = "INFORMACI�N COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE ENV�O: " & strHojaEnvio & vbCrLf & _
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
    ' 5. Procesar seg�n el resultado de la comparaci�n
    '--------------------------------------------------------------------------
    
    
    '----------------------------------------------------------------------
    ' 5.1. Rangos iguales: Copiar datos espec�ficos (filas+2, columnas+11)
    '----------------------------------------------------------------------
    lngLineaError = 90003
    fun801_LogMessage "Ejecutando copia espec�fica para rangos id�nticos...", False, "", strFuncion
    
    ' Validar que hay suficientes filas y columnas para el offset
    'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
       (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
        
        ' Definir rango origen (desde comprobaci�n)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino (hacia env�o)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores �nicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                "Error al copiar valores espec�ficos"
        End If
        
        fun801_LogMessage "Copia espec�fica completada correctamente", False, "", strFuncion
    'Else
    '    fun801_LogMessage "Advertencia: Offset insuficiente para copia espec�fica, omitiendo operaci�n", False, "", strFuncion
    'End If
    
        
    '--------------------------------------------------------------------------
    ' 6. Verificar integridad de la operaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operaci�n...", False, "", strFuncion
    
    ' Verificaci�n b�sica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificaci�n completada: ambas hojas est�n vac�as (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificaci�n", False, "", strFuncion
    Else
        fun801_LogMessage "Verificaci�n completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Comprobar cada celda y etiquetar en color Verde o ROJO cada l�nea
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
    fun801_LogMessage "Copia de datos de comprobaci�n a env�o completada con �xito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F008_Actualizar_Informe_PL_AdHoc = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuraci�n
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F008_Actualizar_Informe_PL_AdHoc = False
End Function

Public Function F007_Preparar_Datos_para_Borrado(ByVal strHojaComprobacion As String, ByVal strHojaEnvio As String) As Boolean
    
    '******************************************************************************
    ' FUNCI�N PRINCIPAL: F007_Preparar_Datos_para_Borrado
    ' Fecha y Hora de Creaci�n: 2025-06-03 00:14:44 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Copia datos espec�ficos desde la hoja de comprobaci�n hacia la hoja de env�o,
    ' implementando l�gica condicional basada en la comparaci�n de rangos entre ambas hojas.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Validar par�metros y obtener referencias a hojas de trabajo
    ' 2. Detectar rangos de datos en hoja de comprobaci�n
    ' 3. Detectar rangos de datos en hoja de env�o
    ' 4. Comparar si los rangos son id�nticos
    ' 5. Si rangos son iguales: copiar datos espec�ficos (filas+2, columnas+11)
    ' 6. Si rangos son diferentes: copiar contenido completo y limpiar excesos
    ' 7. Verificar integridad de la operaci�n
    ' 8. Registrar resultado exitoso en el log del sistema
    '
    ' Par�metros:
    ' - strHojaEnvio: Nombre de la hoja de destino (env�o)
    ' - strHojaComprobacion: Nombre de la hoja de origen (comprobaci�n)
    '
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    'Variable para habilitar/deshabilitar partes de esta funcion
    Dim vEnabled_Parts As Boolean
    
    ' Variables para mostrar informaci�n de rangos
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
    
    ' Variables para rangos de la hoja de comprobaci�n
    Dim vFila_Inicial_HojaComprobacion As Long
    Dim vFila_Final_HojaComprobacion As Long
    Dim vColumna_Inicial_HojaComprobacion As Long
    Dim vColumna_Final_HojaComprobacion As Long
    
    ' Variables para rangos de la hoja de env�o
    Dim vFila_Inicial_HojaEnvio As Long
    Dim vFila_Final_HojaEnvio As Long
    Dim vColumna_Inicial_HojaEnvio As Long
    Dim vColumna_Final_HojaEnvio As Long
    
    ' Variable para comparaci�n de rangos
    Dim vLosRangosSonIguales As Boolean
    
    ' Variables para rangos de copia
    Dim rngOrigen As Range
    Dim rngDestino As Range
    
    ' Inicializaci�n
    strFuncion = "F007_Preparar_Datos_para_Borrado" 'La funcion Caller es valida solo desde Excel 2000, para Excel 97 usariamos: strFuncion = "F007_Preparar_Datos_para_Borrado"
    F007_Preparar_Datos_para_Borrado = False
    lngLineaError = 0
    vLosRangosSonIguales = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Validar par�metros y obtener referencias a hojas de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Validando hojas para borrado de datos antig�os ...", False, "", strFuncion
    
    ' Validar hoja de env�o
    If Not fun802_SheetExists(strHojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 701, strFuncion, _
            "La hoja de env�o no existe: " & strHojaEnvio
    End If
    
    ' Validar hoja de comprobaci�n
    If Not fun802_SheetExists(strHojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 702, strFuncion, _
            "La hoja de comprobaci�n no existe: " & strHojaComprobacion
    End If
    
    ' Obtener referencias a las hojas
    Set wsEnvio = ThisWorkbook.Worksheets(strHojaEnvio)
    Set wsComprobacion = ThisWorkbook.Worksheets(strHojaComprobacion)
    
    ' Verificar que las referencias son v�lidas
    If wsEnvio Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 703, strFuncion, _
            "No se pudo obtener referencia a la hoja de env�o"
    End If
    
    If wsComprobacion Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 704, strFuncion, _
            "No se pudo obtener referencia a la hoja de comprobaci�n"
    End If
    
    '--------------------------------------------------------------------------
    ' 2. OPCIONAL: Configurar palabras clave espec�ficas si es necesario
    '--------------------------------------------------------------------------
    lngLineaError = 55
    ' Configurar palabras clave para este procesamiento espec�fico
    ' Solo si necesitas valores diferentes a los por defecto
    Dim vEscenarioAdmitido, vUltimoMesCarga As String
    vEscenarioAdmitido = UCase(Trim(CONST_ESCENARIO_ADMITIDO))
    vUltimoMesCarga = UCase(Trim(CONST_ULTIMO_MES_DE_CARGA))
    Call fun826_ConfigurarPalabrasClave(vEscenarioAdmitido, vEscenarioAdmitido, vEscenarioAdmitido, vUltimoMesCarga)
    
    '--------------------------------------------------------------------------
    ' 2. Detectar rangos de datos en hoja de comprobaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 60
    fun801_LogMessage "Detectando rangos de datos en hoja de comprobaci�n...", False, "", strHojaComprobacion
    
    If Not fun822_DetectarRangoCompletoHoja(wsComprobacion, _
                                           vFila_Inicial_HojaComprobacion, _
                                           vFila_Final_HojaComprobacion, _
                                           vColumna_Inicial_HojaComprobacion, _
                                           vColumna_Final_HojaComprobacion) Then
        Err.Raise ERROR_BASE_IMPORT + 705, strFuncion, _
            "Error al detectar rangos en hoja de comprobaci�n"
    End If
    
    fun801_LogMessage "Rangos de comprobaci�n - Filas: " & vFila_Inicial_HojaComprobacion & " a " & vFila_Final_HojaComprobacion & _
                      ", Columnas: " & vColumna_Inicial_HojaComprobacion & " a " & vColumna_Final_HojaComprobacion, _
                      False, "", strHojaComprobacion
    
    vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaComprobacion - 1 'Le quitamos 1, para que considere tambi�n la fila en la que est�n los headers de los meses M01 ... M12
    vColumna_Final_HojaComprobacion = vColumna_Inicial_HojaComprobacion + 22
    
    '--------------------------------------------------------------------------
    ' 3. Detectar rangos de datos en hoja de env�o
    '--------------------------------------------------------------------------
    lngLineaError = 70
    fun801_LogMessage "Detectando rangos de datos en hoja de env�o...", False, "", strHojaEnvio
    
    If Not fun822_DetectarRangoCompletoHoja(wsEnvio, _
                                           vFila_Inicial_HojaEnvio, _
                                           vFila_Final_HojaEnvio, _
                                           vColumna_Inicial_HojaEnvio, _
                                           vColumna_Final_HojaEnvio) Then
        Err.Raise ERROR_BASE_IMPORT + 706, strFuncion, _
            "Error al detectar rangos en hoja de env�o"
    End If
    
    fun801_LogMessage "Rangos de env�o - Filas: " & vFila_Inicial_HojaEnvio & " a " & vFila_Final_HojaEnvio & _
                      ", Columnas: " & vColumna_Inicial_HojaEnvio & " a " & vColumna_Final_HojaEnvio, _
                      False, "", strHojaEnvio
            
    vFila_Inicial_HojaEnvio = vFila_Inicial_HojaEnvio - 1 'Le quitamos 1, para que considere tambi�n la fila en la que est�n los headers de los meses M01 ... M12
    vColumna_Final_HojaEnvio = vColumna_Inicial_HojaEnvio + 22
            
    '--------------------------------------------------------------------------
    ' 3.1. NUEVO: Mostrar informaci�n completa de rangos de ambas hojas
    '--------------------------------------------------------------------------
    
    vEnabled_Parts = False
    If vEnabled_Parts Then

        lngLineaError = 125
        strMensajeRangosCompleto = "INFORMACI�N COMPLETA DE RANGOS DETECTADOS" & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE ENV�O: " & strHojaEnvio & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "- Fila Inicial: " & vFila_Inicial_HojaEnvio & vbCrLf & _
                                   "- Fila Final: " & vFila_Final_HojaEnvio & vbCrLf & _
                                   "- Columna Inicial: " & vColumna_Inicial_HojaEnvio & vbCrLf & _
                                   "- Columna Final: " & vColumna_Final_HojaEnvio & vbCrLf & _
                                   "- Total filas: " & (vFila_Final_HojaEnvio - vFila_Inicial_HojaEnvio + 1) & vbCrLf & _
                                   "- Total columnas: " & (vColumna_Final_HojaEnvio - vColumna_Inicial_HojaEnvio + 1) & vbCrLf & vbCrLf & _
                                   "-----------------------------------------------" & vbCrLf & _
                                   "HOJA DE COMPROBACI�N: " & strHojaComprobacion & vbCrLf & _
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
    ' 4. Comparar si los rangos son id�nticos
    '--------------------------------------------------------------------------
    lngLineaError = 80
    fun801_LogMessage "Comparando rangos entre hojas...", False, "", strFuncion
    
    If (vFila_Inicial_HojaComprobacion = vFila_Inicial_HojaEnvio) And _
       (vFila_Final_HojaComprobacion = vFila_Final_HojaEnvio) And _
       (vColumna_Inicial_HojaComprobacion = vColumna_Inicial_HojaEnvio) And _
       (vColumna_Final_HojaComprobacion = vColumna_Final_HojaEnvio) Then
        vLosRangosSonIguales = True
        fun801_LogMessage "Los rangos son id�nticos - Aplicando copia espec�fica", False, "", strFuncion
    Else
        vLosRangosSonIguales = False
        fun801_LogMessage "Los rangos son diferentes - Aplicando copia completa", False, "", strFuncion
    End If
    
    'MsgBox "Los Rangos son Iguales? = " & vLosRangosSonIguales
    
    'En realidad si los rangos no salen iguales, tiene que ser
    '   porque en una de las 2 hojas est� considerando como "Contenido"
    '   algunas celdas que en realidad no tienen contenido
    '   (tendr�amos que hacerle un ClearConents a algunos rangos,
    '   como por ejemplo columnas anteriores a la del primer "BUDGET_OS", columnas posteriores a la del "M12"
    '   o filas anteriores a la del M12
    
    'Asi que vamos a forzar a que los rangos sean iguales
    ' y vamos a usar los rangos de la strHojaComprobacion
    vLosRangosSonIguales = True
    
    '--------------------------------------------------------------------------
    ' 5. Procesar seg�n el resultado de la comparaci�n
    '--------------------------------------------------------------------------
    If vLosRangosSonIguales = True Then
        '----------------------------------------------------------------------
        ' 5.1. Rangos iguales: Copiar datos espec�ficos (filas+2, columnas+11)
        '----------------------------------------------------------------------
        lngLineaError = 90003
        fun801_LogMessage "Ejecutando copia espec�fica para rangos id�nticos...", False, "", strFuncion
        
        ' Validar que hay suficientes filas y columnas para el offset
        'If (vFila_Inicial_HojaComprobacion + 2) <= vFila_Final_HojaComprobacion And _
           (vColumna_Inicial_HojaComprobacion + 11) <= vColumna_Final_HojaComprobacion Then
            
            ' Definir rango origen (desde comprobaci�n)
            Set rngOrigen = wsComprobacion.Range( _
                wsComprobacion.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Definir rango destino (hacia env�o)
            Set rngDestino = wsEnvio.Range( _
                wsEnvio.Cells(vFila_Inicial_HojaComprobacion + 2, vColumna_Inicial_HojaComprobacion + 11), _
                wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
            
            ' Realizar copia de valores �nicamente
            If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
                Err.Raise ERROR_BASE_IMPORT + 707, strFuncion, _
                    "Error al copiar valores espec�ficos"
            End If
            
            fun801_LogMessage "Copia espec�fica completada correctamente", False, "", strFuncion
        'Else
        '    fun801_LogMessage "Advertencia: Offset insuficiente para copia espec�fica, omitiendo operaci�n", False, "", strFuncion
        'End If
        
    Else
        '----------------------------------------------------------------------
        ' 5.2. Rangos diferentes: Copiar contenido completo de HojaComprobacion a HojaEnvio
        '----------------------------------------------------------------------
        lngLineaError = 100
        fun801_LogMessage "Ejecutando copia completa para rangos diferentes...", False, "", strFuncion
        
        ' Definir rango origen completo (desde comprobaci�n)
        Set rngOrigen = wsComprobacion.Range( _
            wsComprobacion.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsComprobacion.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Definir rango destino completo (hacia env�o)
        Set rngDestino = wsEnvio.Range( _
            wsEnvio.Cells(vFila_Inicial_HojaComprobacion, vColumna_Inicial_HojaComprobacion), _
            wsEnvio.Cells(vFila_Final_HojaComprobacion, vColumna_Final_HojaComprobacion))
        
        ' Realizar copia de valores �nicamente
        If Not fun823_CopiarSoloValores(rngOrigen, rngDestino) Then
            Err.Raise ERROR_BASE_IMPORT + 708, strFuncion, _
                "Error al copiar contenido completo"
        End If
        
        '----------------------------------------------------------------------
        ' 5.3. Limpiar excesos en hoja de env�o
        '----------------------------------------------------------------------
        lngLineaError = 110
        fun801_LogMessage "Limpiando excesos en hoja de env�o...", False, "", strHojaEnvio
        
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
    ' 6. Verificar integridad de la operaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 120
    fun801_LogMessage "Verificando integridad de la operaci�n...", False, "", strFuncion
    
    ' Verificaci�n b�sica: comprobar que las hojas mantienen contenido coherente
    If wsComprobacion.UsedRange Is Nothing And wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Verificaci�n completada: ambas hojas est�n vac�as (coherente)", False, "", strFuncion
    ElseIf wsComprobacion.UsedRange Is Nothing Or wsEnvio.UsedRange Is Nothing Then
        fun801_LogMessage "Advertencia: Inconsistencia detectada en verificaci�n", False, "", strFuncion
    Else
        fun801_LogMessage "Verificaci�n completada: ambas hojas contienen datos", False, "", strFuncion
    End If
    
    
    '--------------------------------------------------------------------------
    ' 6.1. Editar cada celda para que luego el Submit pueda funcionar
    '--------------------------------------------------------------------------
    lngLineaError = 125
    fun801_LogMessage "Editando cada celda del rango para poder hacer Submit y borrar los datos antig�os...", False, "", strFuncion
    
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
                'Si el escenario no es el "admitido", entonces ponemos un valor "incorrecto" para que no env�e dato
                wsEnvio.Cells(r, vColumna_Inicial_HojaComprobacion).Value = "ESCENARIO_INCORRECTO"
            End If
        Next c
    Next r
    Application.ScreenUpdating = True
    
    '--------------------------------------------------------------------------
    ' 7. Registrar resultado exitoso
    '--------------------------------------------------------------------------
    lngLineaError = 130
    fun801_LogMessage "Copia de datos de comprobaci�n a env�o completada con �xito", _
                      False, strHojaComprobacion, strHojaEnvio
    
    F007_Preparar_Datos_para_Borrado = True
    Exit Function

GestorErrores:
    ' Limpiar objetos y restaurar configuraci�n
    Application.CutCopyMode = False
    Set rngOrigen = Nothing
    Set rngDestino = Nothing
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    MsgBox strMensajeError, vbCritical, "Error en " & strFuncion
    F007_Preparar_Datos_para_Borrado = False
End Function


'Public Function F009_Localizar_Hoja_Envio_Anterior(ByVal vScenario_HEnvio As String, ByVal vYear_HEnvio As String, ByVal vEntity_HEnvio As String) As Boolean
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


