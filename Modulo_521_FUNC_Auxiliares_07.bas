Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_07"
Option Explicit

Public Function fun801_VerificarExistenciaHoja(ByRef wb As Workbook, ByVal strNombreHoja As String) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 801: VERIFICAR EXISTENCIA DE HOJA
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:27:06 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Verifica si existe una hoja espec�fica en un libro de trabajo
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Inicializaci�n
    strFuncion = "fun801_VerificarExistenciaHoja"
    fun801_VerificarExistenciaHoja = False
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: VALIDAR PAR�METROS DE ENTRADA
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Validando par�metros de entrada"
    fun801_LogMessage "[INICIO] " & strFuncion & " - Hoja buscada: '" & strNombreHoja & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If wb Is Nothing Then
        strMensajeError = "El libro de trabajo proporcionado es Nothing"
        Err.Raise ERROR_BASE_IMPORT + 1004, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[DETALLE] Libro v�lido - Nombre: '" & wb.Name & "' | Ruta: '" & wb.Path & "' | Total hojas: " & wb.Worksheets.Count, False, "", strFuncion
    
    If Len(Trim(strNombreHoja)) = 0 Then
        strMensajeError = "El nombre de la hoja no puede estar vac�o"
        Err.Raise ERROR_BASE_IMPORT + 1005, strFuncion, strMensajeError
    End If
    
    lngLineaError = 110
    
    '--------------------------------------------------------------------------
    ' PASO 2: RECORRER TODAS LAS HOJAS DEL LIBRO
    '--------------------------------------------------------------------------
    lngLineaError = 120
    strContexto = "Recorriendo hojas del libro para b�squeda case-insensitive"
    fun801_LogMessage "[PROCESO] " & strContexto & " - Buscando: '" & UCase(Trim(strNombreHoja)) & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    For i = 1 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        
        fun801_LogMessage "[ITERACI�N " & i & "] Comparando hoja: '" & ws.Name & "' vs '" & strNombreHoja & "' (UCase)", False, "", strFuncion
        
        '----------------------------------------------------------------------
        ' PASO 3: COMPARAR NOMBRES DE FORMA CASE-INSENSITIVE
        '----------------------------------------------------------------------
        If UCase(Trim(ws.Name)) = UCase(Trim(strNombreHoja)) Then
            fun801_LogMessage "[�XITO] Hoja encontrada en posici�n " & i & " - Nombre exacto: '" & ws.Name & "' | Estado visible: " & ws.Visible, False, "", strFuncion
            fun801_VerificarExistenciaHoja = True
            Set ws = Nothing
            Exit Function
        End If
    Next i
    
    lngLineaError = 130
    
    '--------------------------------------------------------------------------
    ' PASO 4: HOJA NO ENCONTRADA
    '--------------------------------------------------------------------------
    strContexto = "Finalizando b�squeda - hoja no encontrada"
    fun801_LogMessage "[RESULTADO] " & strContexto & " - '" & strNombreHoja & "' no existe en el libro '" & wb.Name & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ' Listar todas las hojas disponibles para debugging
    Dim strListaHojas As String
    strListaHojas = ""
    For i = 1 To wb.Worksheets.Count
        strListaHojas = strListaHojas & "[" & i & "]" & wb.Worksheets(i).Name
        If i < wb.Worksheets.Count Then strListaHojas = strListaHojas & " | "
    Next i
    fun801_LogMessage "[DEBUG] Hojas disponibles en '" & wb.Name & "': " & strListaHojas, False, "", strFuncion
    
    fun801_VerificarExistenciaHoja = False
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Hoja buscada: '" & strNombreHoja & "'" & vbCrLf & _
                      "Libro: " & IIf(wb Is Nothing, "Nothing", wb.Name) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun801_VerificarExistenciaHoja = False
    
    Set ws = Nothing
End Function

Public Function fun802_CrearHojaDelimitadores(ByRef wb As Workbook, ByVal strNombreHoja As String) As Worksheet
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 802: CREAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:27:06 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Crea una nueva hoja para almacenar delimitadores originales
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim ws As Worksheet
    Dim blnScreenUpdating As Boolean
    Dim intHojasAntes As Integer
    Dim intHojasDespues As Integer
    
    ' Inicializaci�n
    strFuncion = "fun802_CrearHojaDelimitadores"
    Set fun802_CrearHojaDelimitadores = Nothing
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: VALIDAR PAR�METROS DE ENTRADA Y ESTADO INICIAL
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Validando par�metros y estado inicial del libro"
    fun801_LogMessage "[INICIO] " & strFuncion & " - Creando hoja: '" & strNombreHoja & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If wb Is Nothing Then
        strMensajeError = "El libro de trabajo proporcionado es Nothing"
        Err.Raise ERROR_BASE_IMPORT + 1006, strFuncion, strMensajeError
    End If
    
    If Len(Trim(strNombreHoja)) = 0 Then
        strMensajeError = "El nombre de la hoja no puede estar vac�o"
        Err.Raise ERROR_BASE_IMPORT + 1007, strFuncion, strMensajeError
    End If
    
    intHojasAntes = wb.Worksheets.Count
    fun801_LogMessage "[DETALLE] Estado inicial - Libro: '" & wb.Name & "' | Hojas existentes: " & intHojasAntes & _
                      " | Protegido: " & wb.ProtectStructure, False, "", strFuncion
    
    lngLineaError = 110
    
    '--------------------------------------------------------------------------
    ' PASO 2: VERIFICAR QUE NO EXISTA YA LA HOJA
    '--------------------------------------------------------------------------
    lngLineaError = 120
    strContexto = "Verificando si la hoja ya existe antes de crear"
    fun801_LogMessage "[PASO 1] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If fun801_VerificarExistenciaHoja(wb, strNombreHoja) Then
        fun801_LogMessage "[ADVERTENCIA] La hoja ya existe, retornando referencia a hoja existente: '" & strNombreHoja & "'", False, "", strFuncion
        Set fun802_CrearHojaDelimitadores = wb.Worksheets(strNombreHoja)
        Exit Function
    End If
    
    lngLineaError = 130
    
    '--------------------------------------------------------------------------
    ' PASO 3: VERIFICAR PERMISOS Y PREPARAR ENTORNO
    '--------------------------------------------------------------------------
    lngLineaError = 140
    strContexto = "Verificando permisos y preparando entorno de trabajo"
    fun801_LogMessage "[PASO 2] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If wb.ProtectStructure Then
        strMensajeError = "No se puede crear hoja: el libro est� protegido contra cambios estructurales"
        Err.Raise ERROR_BASE_IMPORT + 1008, strFuncion, strMensajeError
    End If
    
    ' Optimizaci�n: deshabilitar actualizaciones
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    fun801_LogMessage "[OPTIMIZACI�N] ScreenUpdating deshabilitado para mejorar rendimiento", False, "", strFuncion
    
    lngLineaError = 150
    
    '--------------------------------------------------------------------------
    ' PASO 4: CREAR NUEVA HOJA AL FINAL DEL LIBRO
    '--------------------------------------------------------------------------
    lngLineaError = 160
    strContexto = "Creando nueva hoja al final del libro"
    fun801_LogMessage "[PASO 3] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    fun801_LogMessage "[�XITO] Hoja base creada en posici�n: " & ws.Index, False, "", strFuncion
    
    lngLineaError = 170
    
    '--------------------------------------------------------------------------
    ' PASO 5: ASIGNAR NOMBRE Y CONFIGURAR PROPIEDADES
    '--------------------------------------------------------------------------
    lngLineaError = 180
    strContexto = "Asignando nombre y configurando propiedades b�sicas"
    fun801_LogMessage "[PASO 4] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ws.Name = strNombreHoja
    fun801_LogMessage "[DETALLE] Nombre asignado exitosamente: '" & ws.Name & "'", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 6: CONFIGURAR FORMATO Y PROPIEDADES AVANZADAS
    '--------------------------------------------------------------------------
    lngLineaError = 190
    strContexto = "Configurando formato y propiedades de la hoja"
    fun801_LogMessage "[PASO 5] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    With ws
        .DisplayGridlines = True
        .DisplayHeadings = True
        .Range("A1").Select
        .Columns.StandardWidth = 10
        .PageSetup.PrintArea = ""
        .Visible = xlSheetVisible
        
        ' Verificar y desproteger si es necesario
        If .ProtectContents Then
            .Unprotect
            fun801_LogMessage "[DETALLE] Hoja desprotegida para permitir modificaciones", False, "", strFuncion
        End If
    End With
    
    lngLineaError = 200
    
    '--------------------------------------------------------------------------
    ' PASO 7: VERIFICACI�N FINAL Y LIMPIEZA
    '--------------------------------------------------------------------------
    lngLineaError = 210
    strContexto = "Verificaci�n final y restauraci�n del entorno"
    fun801_LogMessage "[PASO 6] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    intHojasDespues = wb.Worksheets.Count
    fun801_LogMessage "[VERIFICACI�N] Hojas antes: " & intHojasAntes & " | Hojas despu�s: " & intHojasDespues & _
                      " | Diferencia: " & (intHojasDespues - intHojasAntes), False, "", strFuncion
    
    ' Restaurar configuraci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdating
    fun801_LogMessage "[OPTIMIZACI�N] ScreenUpdating restaurado", False, "", strFuncion
    
    lngLineaError = 220
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Hoja: '" & ws.Name & "' | �ndice: " & ws.Index, False, "", strFuncion
    Set fun802_CrearHojaDelimitadores = ws
    
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
                      "Hoja a crear: '" & strNombreHoja & "'" & vbCrLf & _
                      "Libro: " & IIf(wb Is Nothing, "Nothing", wb.Name) & vbCrLf & _
                      "Hojas antes del error: " & intHojasAntes & vbCrLf & _
                      "Libro protegido: " & IIf(wb Is Nothing, "N/A", CStr(wb.ProtectStructure)) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    Set fun802_CrearHojaDelimitadores = Nothing
    
    Set ws = Nothing
End Function

Public Function fun803_HacerHojaVisible(ByRef ws As Worksheet) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 803: HACER HOJA VISIBLE
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:27:06 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Hace visible una hoja si est� oculta
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim estadoAnterior As Integer
    Dim estadoNuevo As Integer
    
    ' Inicializaci�n
    strFuncion = "fun803_HacerHojaVisible"
    fun803_HacerHojaVisible = False
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: VALIDAR PAR�METRO DE ENTRADA Y ESTADO INICIAL
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Validando par�metro de entrada y detectando estado inicial"
    
    If ws Is Nothing Then
        strMensajeError = "La hoja de trabajo proporcionada es Nothing"
        Err.Raise ERROR_BASE_IMPORT + 1009, strFuncion, strMensajeError
    End If
    
    estadoAnterior = ws.Visible
    fun801_LogMessage "[INICIO] " & strFuncion & " - Hoja: '" & ws.Name & "' | Estado inicial: " & estadoAnterior & _
                      " (xlSheetVisible=-1, xlSheetHidden=0, xlSheetVeryHidden=2) (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    lngLineaError = 110
    
    '--------------------------------------------------------------------------
    ' PASO 2: VERIFICAR PERMISOS DE MODIFICACI�N
    '--------------------------------------------------------------------------
    lngLineaError = 120
    strContexto = "Verificando permisos para cambiar visibilidad"
    fun801_LogMessage "[PASO 1] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If ws.Parent.ProtectStructure Then
        fun801_LogMessage "[ADVERTENCIA] No se puede cambiar visibilidad: libro protegido - Libro: '" & ws.Parent.Name & "'", False, "", strFuncion
        fun803_HacerHojaVisible = False
        Exit Function
    End If
    
    lngLineaError = 130
    
    '--------------------------------------------------------------------------
    ' PASO 3: ANALIZAR ESTADO ACTUAL Y DETERMINAR ACCI�N
    '--------------------------------------------------------------------------
    lngLineaError = 140
    strContexto = "Analizando estado actual de visibilidad"
    fun801_LogMessage "[PASO 2] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    Select Case estadoAnterior
        Case xlSheetVisible
            fun801_LogMessage "[RESULTADO] La hoja ya est� visible - No se requiere acci�n", False, "", strFuncion
            fun803_HacerHojaVisible = True
            Exit Function
            
        Case xlSheetHidden
            fun801_LogMessage "[ACCI�N] Hoja oculta (xlSheetHidden) - Procediendo a hacer visible", False, "", strFuncion
            
        Case xlSheetVeryHidden
            fun801_LogMessage "[ACCI�N] Hoja muy oculta (xlSheetVeryHidden) - Procediendo a hacer visible", False, "", strFuncion
            
        Case Else
            fun801_LogMessage "[ACCI�N] Estado desconocido (" & estadoAnterior & ") - Forzando visibilidad", False, "", strFuncion
    End Select
    
    lngLineaError = 150
    
    '--------------------------------------------------------------------------
    ' PASO 4: APLICAR CAMBIO DE VISIBILIDAD
    '--------------------------------------------------------------------------
    lngLineaError = 160
    strContexto = "Aplicando cambio de visibilidad"
    fun801_LogMessage "[PASO 3] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ws.Visible = xlSheetVisible
    estadoNuevo = ws.Visible
    
    lngLineaError = 170
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR QUE EL CAMBIO SE APLIC� CORRECTAMENTE
    '--------------------------------------------------------------------------
    lngLineaError = 180
    strContexto = "Verificando resultado del cambio de visibilidad"
    fun801_LogMessage "[PASO 4] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If estadoNuevo = xlSheetVisible Then
        fun801_LogMessage "[�XITO] Cambio de visibilidad exitoso - Estado anterior: " & estadoAnterior & " | Estado nuevo: " & estadoNuevo, False, "", strFuncion
        fun803_HacerHojaVisible = True
    Else
        strMensajeError = "Fallo en cambio de visibilidad - Estado anterior: " & estadoAnterior & " | Estado actual: " & estadoNuevo & " | Esperado: " & xlSheetVisible
        Err.Raise ERROR_BASE_IMPORT + 1010, strFuncion, strMensajeError
    End If
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Hoja: " & IIf(ws Is Nothing, "Nothing", ws.Name) & vbCrLf & _
                      "Estado anterior: " & estadoAnterior & vbCrLf & _
                      "Estado nuevo: " & estadoNuevo & vbCrLf & _
                      "Libro protegido: " & IIf(ws Is Nothing, "N/A", CStr(ws.Parent.ProtectStructure)) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun803_HacerHojaVisible = False
End Function

Public Function fun804_ConvertirValorACadena(ByVal valorCelda As Variant) As String
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 804: CONVERTIR VALOR A CADENA
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:27:06 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Convierte un valor de celda a cadena de forma segura
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim tipoValor As String
    Dim valorOriginal As String
    
    ' Inicializaci�n
    strFuncion = "fun804_ConvertirValorACadena"
    fun804_ConvertirValorACadena = ""
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: AN�LISIS INICIAL DEL TIPO DE VALOR
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Analizando tipo y estado del valor de entrada"
    tipoValor = TypeName(valorCelda)
    fun801_LogMessage "[INICIO] " & strFuncion & " - Tipo detectado: " & tipoValor & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 2: VERIFICAR SI EL VALOR ES EMPTY O NULL
    '--------------------------------------------------------------------------
    lngLineaError = 110
    If IsEmpty(valorCelda) Then
        strContexto = "Valor Empty detectado"
        fun801_LogMessage "[RESULTADO] " & strContexto & " - Retornando cadena vac�a (L�nea: " & lngLineaError & ")", False, "", strFuncion
        fun804_ConvertirValorACadena = ""
        Exit Function
    End If
    
    lngLineaError = 120
    If IsNull(valorCelda) Then
        strContexto = "Valor Null detectado"
        fun801_LogMessage "[RESULTADO] " & strContexto & " - Retornando cadena vac�a (L�nea: " & lngLineaError & ")", False, "", strFuncion
        fun804_ConvertirValorACadena = ""
        Exit Function
    End If
    
    lngLineaError = 130
    
    '--------------------------------------------------------------------------
    ' PASO 3: VERIFICAR SI EL VALOR ES ERROR
    '--------------------------------------------------------------------------
    lngLineaError = 140
    If IsError(valorCelda) Then
        strContexto = "Error en celda detectado"
        fun801_LogMessage "[ADVERTENCIA] " & strContexto & " - Error: " & CStr(valorCelda) & " - Retornando cadena vac�a (L�nea: " & lngLineaError & ")", False, "", strFuncion
        fun804_ConvertirValorACadena = ""
        Exit Function
    End If
    
    lngLineaError = 150
    
    '--------------------------------------------------------------------------
    ' PASO 4: CONVERTIR A CADENA Y APLICAR TRIM
    '--------------------------------------------------------------------------
    lngLineaError = 160
    strContexto = "Convirtiendo valor a cadena y aplicando Trim"
    
    ' Guardar representaci�n original para logging
    On Error Resume Next
    valorOriginal = CStr(valorCelda)
    If Err.Number <> 0 Then
        fun801_LogMessage "[ERROR] Error en conversi�n CStr - Error: " & Err.Number & " - " & Err.Description & " (L�nea: " & lngLineaError & ")", True, "", strFuncion
        On Error GoTo GestorErrores
        Err.Raise ERROR_BASE_IMPORT + 1011, strFuncion, "Error en conversi�n CStr del valor: " & tipoValor
    End If
    On Error GoTo GestorErrores
    
    fun804_ConvertirValorACadena = Trim(valorOriginal)
    
    fun801_LogMessage "[PROCESO] " & strContexto & " - Valor original: '" & valorOriginal & "' | Longitud original: " & Len(valorOriginal) & _
                      " | Valor trimmed: '" & fun804_ConvertirValorACadena & "' | Longitud final: " & Len(fun804_ConvertirValorACadena) & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    lngLineaError = 170
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Resultado: '" & fun804_ConvertirValorACadena & "'", False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Tipo de valor: " & tipoValor & vbCrLf & _
                      "Valor original: " & valorOriginal & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun804_ConvertirValorACadena = ""
End Function

Public Function fun805_ValidarValoresOriginales() As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 805: VALIDAR VALORES ORIGINALES
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 23:12:13 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Valida que los valores originales le�dos sean apropiados para restauraci�n
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim valoresValidos As Integer
    Dim detalleValidacion As String
    Dim valorUseSystem As String
    Dim valorDecimal As String
    Dim valorThousands As String
    
    ' Inicializaci�n
    strFuncion = "fun805_ValidarValoresOriginales"
    fun805_ValidarValoresOriginales = False
    lngLineaError = 0
    valoresValidos = 0
    strContexto = ""
    detalleValidacion = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y PREPARACI�N
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando validaci�n de valores originales para restauraci�n"
    fun801_LogMessage "[INICIO] " & strFuncion & " - " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    fun801_LogMessage "[VALORES A VALIDAR] UseSystem: '" & vExcel_UseSystemSeparators_ValorOriginal & _
                      "' | Decimal: '" & vExcel_DecimalSeparator_ValorOriginal & _
                      "' | Thousands: '" & vExcel_ThousandsSeparator_ValorOriginal & "'", False, "", strFuncion
    
    ' Normalizar valores para an�lisis
    valorUseSystem = UCase(Trim(vExcel_UseSystemSeparators_ValorOriginal))
    valorDecimal = Trim(vExcel_DecimalSeparator_ValorOriginal)
    valorThousands = Trim(vExcel_ThousandsSeparator_ValorOriginal)
    
    fun801_LogMessage "[NORMALIZACI�N] UseSystem normalizado: '" & valorUseSystem & _
                      "' | Decimal normalizado: '" & valorDecimal & _
                      "' | Thousands normalizado: '" & valorThousands & "'", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 2: VALIDAR USESYSTEMSEPARATORS
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Validando UseSystemSeparators"
    fun801_LogMessage "[PASO 1] " & strContexto & " - Valor: '" & vExcel_UseSystemSeparators_ValorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    fun801_LogMessage "[DETALLE] Longitud original: " & Len(vExcel_UseSystemSeparators_ValorOriginal) & _
                      " | Valor normalizado: '" & valorUseSystem & "' | Longitud normalizada: " & Len(valorUseSystem), False, "", strFuncion
    
    If valorUseSystem = "TRUE" Or valorUseSystem = "FALSE" Then
        valoresValidos = valoresValidos + 1
        fun801_LogMessage "[�XITO] UseSystemSeparators es v�lido: '" & vExcel_UseSystemSeparators_ValorOriginal & "' (normalizado: '" & valorUseSystem & "')", False, "", strFuncion
        detalleValidacion = detalleValidacion & "UseSystemSeparators: V�LIDO (" & valorUseSystem & ") | "
    Else
        fun801_LogMessage "[FALLO] UseSystemSeparators no es v�lido - Esperado: 'True' o 'False' | Recibido: '" & vExcel_UseSystemSeparators_ValorOriginal & _
                          "' | Normalizado: '" & valorUseSystem & "' | Longitud: " & Len(vExcel_UseSystemSeparators_ValorOriginal), False, "", strFuncion
        detalleValidacion = detalleValidacion & "UseSystemSeparators: INV�LIDO ('" & valorUseSystem & "') | "
    End If
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: VALIDAR DECIMALSEPARATOR
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Validando DecimalSeparator"
    fun801_LogMessage "[PASO 2] " & strContexto & " - Valor: '" & vExcel_DecimalSeparator_ValorOriginal & "' | Longitud: " & Len(vExcel_DecimalSeparator_ValorOriginal) & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(valorDecimal) > 0 Then
        fun801_LogMessage "[DETALLE] Car�cter decimal - ASCII: " & Asc(Left(valorDecimal, 1)) & " | Hex: " & Hex(Asc(Left(valorDecimal, 1))), False, "", strFuncion
    End If
    
    If Len(valorDecimal) = 1 And (valorDecimal = "." Or valorDecimal = ",") Then
        valoresValidos = valoresValidos + 1
        fun801_LogMessage "[�XITO] DecimalSeparator es v�lido: '" & vExcel_DecimalSeparator_ValorOriginal & "' | ASCII: " & Asc(valorDecimal), False, "", strFuncion
        detalleValidacion = detalleValidacion & "DecimalSeparator: V�LIDO ('" & valorDecimal & "') | "
    Else
        fun801_LogMessage "[FALLO] DecimalSeparator no es v�lido - Esperado: '.' o ',' (1 car�cter) | Recibido: '" & vExcel_DecimalSeparator_ValorOriginal & _
                          "' | Normalizado: '" & valorDecimal & "' | Longitud: " & Len(vExcel_DecimalSeparator_ValorOriginal) & _
                          " | ASCII primer car�cter: " & IIf(Len(valorDecimal) > 0, CStr(Asc(Left(valorDecimal, 1))), "N/A"), False, "", strFuncion
        detalleValidacion = detalleValidacion & "DecimalSeparator: INV�LIDO ('" & valorDecimal & "', Long:" & Len(valorDecimal) & ") | "
    End If
    
    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: VALIDAR THOUSANDSSEPARATOR
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Validando ThousandsSeparator"
    fun801_LogMessage "[PASO 3] " & strContexto & " - Valor: '" & vExcel_ThousandsSeparator_ValorOriginal & "' | Longitud: " & Len(vExcel_ThousandsSeparator_ValorOriginal) & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(valorThousands) > 0 Then
        fun801_LogMessage "[DETALLE] Car�cter thousands - ASCII: " & Asc(Left(valorThousands, 1)) & " | Hex: " & Hex(Asc(Left(valorThousands, 1))) & _
                          " | Es espacio: " & IIf(Left(valorThousands, 1) = " ", "S�", "NO") & _
                          " | Es comilla: " & IIf(Left(valorThousands, 1) = Chr(39), "S�", "NO"), False, "", strFuncion
    End If
    
    If Len(valorThousands) = 1 And _
       (valorThousands = "." Or valorThousands = "," Or _
        valorThousands = " " Or valorThousands = Chr(39) Or _
        valorThousands = Chr(160)) Then ' Chr(160) = espacio no separable
        valoresValidos = valoresValidos + 1
        fun801_LogMessage "[�XITO] ThousandsSeparator es v�lido: '" & vExcel_ThousandsSeparator_ValorOriginal & "' | ASCII: " & Asc(valorThousands) & _
                          " | Tipo: " & IIf(valorThousands = ".", "PUNTO", IIf(valorThousands = ",", "COMA", _
                          IIf(valorThousands = " " Or valorThousands = Chr(160), "ESPACIO", _
                          IIf(valorThousands = Chr(39), "COMILLA", "OTRO")))), False, "", strFuncion
        detalleValidacion = detalleValidacion & "ThousandsSeparator: V�LIDO ('" & valorThousands & "', ASCII:" & Asc(valorThousands) & ")"
    Else
        fun801_LogMessage "[FALLO] ThousandsSeparator no es v�lido - Esperado: '.', ',', ' ', ''' o Chr(160) (1 car�cter) | Recibido: '" & vExcel_ThousandsSeparator_ValorOriginal & _
                          "' | Normalizado: '" & valorThousands & "' | Longitud: " & Len(vExcel_ThousandsSeparator_ValorOriginal) & _
                          " | ASCII primer car�cter: " & IIf(Len(valorThousands) > 0, CStr(Asc(Left(valorThousands, 1))), "N/A"), False, "", strFuncion
        detalleValidacion = detalleValidacion & "ThousandsSeparator: INV�LIDO ('" & valorThousands & "', Long:" & Len(valorThousands) & _
                           IIf(Len(valorThousands) > 0, ", ASCII:" & Asc(Left(valorThousands, 1)), "") & ")"
    End If
    
    lngLineaError = 160
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR COMPATIBILIDAD ENTRE DELIMITADORES
    '--------------------------------------------------------------------------
    lngLineaError = 170
    strContexto = "Verificando compatibilidad entre delimitadores"
    fun801_LogMessage "[PASO 4] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If valoresValidos >= 2 And Len(valorDecimal) = 1 And Len(valorThousands) = 1 Then
        If valorDecimal = valorThousands Then
            fun801_LogMessage "[ADVERTENCIA] Decimal y Thousands separators son iguales ('" & valorDecimal & "') - Esto puede causar problemas de formateo", False, "", strFuncion
            detalleValidacion = detalleValidacion & " | ADVERTENCIA: Separadores iguales"
        Else
            fun801_LogMessage "[VERIFICACI�N] Separadores son diferentes - Decimal: '" & valorDecimal & "' | Thousands: '" & valorThousands & "' - Configuraci�n v�lida", False, "", strFuncion
            detalleValidacion = detalleValidacion & " | Separadores diferentes: OK"
        End If
    End If
    
    lngLineaError = 180
    
    '--------------------------------------------------------------------------
    ' PASO 6: EVALUACI�N FINAL DE VALIDACI�N
    '--------------------------------------------------------------------------
    lngLineaError = 190
    strContexto = "Evaluando resultado final de validaci�n"
    fun801_LogMessage "[PASO 5] " & strContexto & " - Valores v�lidos encontrados: " & valoresValidos & "/3 (L�nea: " & lngLineaError & ")", False, "", strFuncion
    fun801_LogMessage "[RESUMEN COMPLETO] " & detalleValidacion, False, "", strFuncion
    
    ' Criterio de validaci�n: Al menos 2 valores v�lidos de 3
    If valoresValidos >= 2 Then
        fun801_LogMessage "[�XITO] Validaci�n exitosa - Suficientes valores v�lidos (" & valoresValidos & "/3) para continuar con restauraci�n", False, "", strFuncion
        fun801_LogMessage "[DECISI�N] Procediendo con restauraci�n usando valores disponibles", False, "", strFuncion
        fun805_ValidarValoresOriginales = True
    ElseIf valoresValidos = 1 Then
        fun801_LogMessage "[ADVERTENCIA] Solo un valor v�lido encontrado (" & valoresValidos & "/3) - Restauraci�n parcial posible pero riesgosa", False, "", strFuncion
        fun801_LogMessage "[DECISI�N] Permitiendo restauraci�n con valores limitados", False, "", strFuncion
        fun805_ValidarValoresOriginales = True
    Else
        strMensajeError = "No se encontraron valores v�lidos para restaurar delimitadores - Todos los valores son inv�lidos: " & detalleValidacion & _
                         " | UseSystem: '" & vExcel_UseSystemSeparators_ValorOriginal & _
                         "' | Decimal: '" & vExcel_DecimalSeparator_ValorOriginal & _
                         "' | Thousands: '" & vExcel_ThousandsSeparator_ValorOriginal & "'"
        fun801_LogMessage "[FALLO CR�TICO] " & strMensajeError, True, "", strFuncion
        Err.Raise ERROR_BASE_IMPORT + 1012, strFuncion, strMensajeError
    End If
    
    lngLineaError = 200
    
    '--------------------------------------------------------------------------
    ' PASO 7: LOGGING FINAL DE VALORES A UTILIZAR
    '--------------------------------------------------------------------------
    fun801_LogMessage "[VALORES FINALES PARA RESTAURACI�N]:", False, "", strFuncion
    fun801_LogMessage "  - UseSystemSeparators: '" & vExcel_UseSystemSeparators_ValorOriginal & "' (v�lido: " & IIf(valorUseSystem = "TRUE" Or valorUseSystem = "FALSE", "S�", "NO") & ")", False, "", strFuncion
    fun801_LogMessage "  - DecimalSeparator: '" & vExcel_DecimalSeparator_ValorOriginal & "' (v�lido: " & IIf(Len(valorDecimal) = 1 And (valorDecimal = "." Or valorDecimal = ","), "S�", "NO") & ")", False, "", strFuncion
    fun801_LogMessage "  - ThousandsSeparator: '" & vExcel_ThousandsSeparator_ValorOriginal & "' (v�lido: " & IIf(Len(valorThousands) = 1, "S�", "NO") & ")", False, "", strFuncion
    
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Validaci�n: APROBADA", False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Valores analizados:" & vbCrLf & _
                      "  - UseSystemSeparators original: '" & vExcel_UseSystemSeparators_ValorOriginal & "'" & vbCrLf & _
                      "  - UseSystemSeparators normalizado: '" & valorUseSystem & "'" & vbCrLf & _
                      "  - DecimalSeparator original: '" & vExcel_DecimalSeparator_ValorOriginal & "'" & vbCrLf & _
                      "  - DecimalSeparator normalizado: '" & valorDecimal & "'" & vbCrLf & _
                      "  - ThousandsSeparator original: '" & vExcel_ThousandsSeparator_ValorOriginal & "'" & vbCrLf & _
                      "  - ThousandsSeparator normalizado: '" & valorThousands & "'" & vbCrLf & _
                      "Valores v�lidos encontrados: " & valoresValidos & "/3" & vbCrLf & _
                      "Detalle validaci�n: " & detalleValidacion & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun805_ValidarValoresOriginales = False
End Function

Public Function fun806_RestaurarUseSystemSeparators(ByVal valorOriginal As String) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 806: RESTAURAR USE SYSTEM SEPARATORS
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:33:04 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura la configuraci�n de UseSystemSeparators
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim valorBooleano As Boolean
    Dim valorActualAntes As Boolean
    Dim valorActualDespues As Boolean
    Dim versionExcel As String
    
    ' Inicializaci�n
    strFuncion = "fun806_RestaurarUseSystemSeparators"
    fun806_RestaurarUseSystemSeparators = False
    lngLineaError = 0
    strContexto = ""
    versionExcel = Application.Version
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y AN�LISIS DEL VALOR
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando restauraci�n de UseSystemSeparators"
    fun801_LogMessage "[INICIO] " & strFuncion & " - Valor a restaurar: '" & valorOriginal & "' | Versi�n Excel: " & versionExcel & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    ' Obtener valor actual antes del cambio
    #If VBA7 Then
        valorActualAntes = Application.UseSystemSeparators
        fun801_LogMessage "[DETALLE] Valor actual antes del cambio: " & valorActualAntes & " (VBA7 disponible)", False, "", strFuncion
    #Else
        fun801_LogMessage "[DETALLE] UseSystemSeparators no disponible en esta versi�n de Excel (VBA6 o anterior)", False, "", strFuncion
        valorActualAntes = False ' Valor por defecto para versiones antiguas
    #End If
    
    '--------------------------------------------------------------------------
    ' PASO 2: VALIDAR PAR�METRO DE ENTRADA
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Validando par�metro de entrada"
    fun801_LogMessage "[PASO 1] " & strContexto & " - Verificando formato del valor (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) = 0 Then
        strMensajeError = "El valor original para UseSystemSeparators no puede estar vac�o"
        Err.Raise ERROR_BASE_IMPORT + 1013, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[DETALLE] Par�metro no vac�o - Longitud: " & Len(valorOriginal) & " | Valor trimmed: '" & Trim(valorOriginal) & "'", False, "", strFuncion
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: CONVERTIR CADENA A VALOR BOOLEANO
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Convirtiendo cadena a valor booleano"
    fun801_LogMessage "[PASO 2] " & strContexto & " - Analizando valor case-insensitive (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If UCase(Trim(valorOriginal)) = "TRUE" Then
        valorBooleano = True
        fun801_LogMessage "[CONVERSI�N] Valor convertido a: True", False, "", strFuncion
    ElseIf UCase(Trim(valorOriginal)) = "FALSE" Then
        valorBooleano = False
        fun801_LogMessage "[CONVERSI�N] Valor convertido a: False", False, "", strFuncion
    Else
        strMensajeError = "Valor no v�lido para UseSystemSeparators - Esperado: 'True' o 'False' | Recibido: '" & valorOriginal & "'"
        Err.Raise ERROR_BASE_IMPORT + 1014, strFuncion, strMensajeError
    End If
    
    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: APLICAR CONFIGURACI�N A EXCEL SEG�N VERSI�N
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Aplicando configuraci�n seg�n versi�n de Excel"
    fun801_LogMessage "[PASO 3] " & strContexto & " - Aplicando UseSystemSeparators = " & valorBooleano & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    #If VBA7 Then
        ' Excel 2010 y posteriores (incluye 365)
        lngLineaError = 160
        fun801_LogMessage "[M�TODO] Usando Application.UseSystemSeparators (VBA7+)", False, "", strFuncion
        Application.UseSystemSeparators = valorBooleano
        fun801_LogMessage "[APLICADO] UseSystemSeparators configurado exitosamente", False, "", strFuncion
    #Else
        ' Excel 97, 2003 y anteriores - usar m�todo alternativo
        lngLineaError = 170
        fun801_LogMessage "[M�TODO] Usando m�todo legacy para versiones anteriores (VBA6)", False, "", strFuncion
        If valorBooleano Then
            fun801_LogMessage "[LEGACY] Configurando delimitadores del sistema (UseSystem=True)", False, "", strFuncion
            Application.DecimalSeparator = Mid(CStr(1.1), 2, 1)
            Application.ThousandsSeparator = ","
        Else
            fun801_LogMessage "[LEGACY] Manteniendo delimitadores personalizados (UseSystem=False)", False, "", strFuncion
            ' No cambiar nada en versiones legacy cuando es False
        End If
    #End If
    
    lngLineaError = 180
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR QUE EL CAMBIO SE APLIC� CORRECTAMENTE
    '--------------------------------------------------------------------------
    lngLineaError = 190
    strContexto = "Verificando aplicaci�n correcta del cambio"
    fun801_LogMessage "[PASO 4] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    #If VBA7 Then
        valorActualDespues = Application.UseSystemSeparators
        fun801_LogMessage "[VERIFICACI�N] Valor despu�s del cambio: " & valorActualDespues & " | Esperado: " & valorBooleano, False, "", strFuncion
        
        If valorActualDespues = valorBooleano Then
            fun801_LogMessage "[�XITO] UseSystemSeparators aplicado correctamente", False, "", strFuncion
            fun806_RestaurarUseSystemSeparators = True
        Else
            strMensajeError = "Error en verificaci�n - Valor esperado: " & valorBooleano & " | Valor actual: " & valorActualDespues
            Err.Raise ERROR_BASE_IMPORT + 1015, strFuncion, strMensajeError
        End If
    #Else
        ' Para versiones anteriores, asumir �xito si no hay error
        fun801_LogMessage "[VERIFICACI�N] M�todo legacy aplicado - AsumIendo �xito (no verificable en VBA6)", False, "", strFuncion
        fun806_RestaurarUseSystemSeparators = True
    #End If
    
    lngLineaError = 200
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Cambio aplicado: " & valorActualAntes & " ? " & valorBooleano, False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Valor Original: '" & valorOriginal & "'" & vbCrLf & _
                      "Valor Booleano: " & valorBooleano & vbCrLf & _
                      "Valor Antes: " & valorActualAntes & vbCrLf & _
                      "Valor Despu�s: " & valorActualDespues & vbCrLf & _
                      "Versi�n Excel: " & versionExcel & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun806_RestaurarUseSystemSeparators = False
End Function

Public Function fun807_RestaurarDecimalSeparator(ByVal valorOriginal As String) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 807: RESTAURAR DECIMAL SEPARATOR
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:33:04 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura la configuraci�n del separador decimal
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim valorActualAntes As String
    Dim valorActualDespues As String
    Dim caracterASCII As Integer
    
    ' Inicializaci�n
    strFuncion = "fun807_RestaurarDecimalSeparator"
    fun807_RestaurarDecimalSeparator = False
    lngLineaError = 0
    strContexto = ""
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y CAPTURA DE ESTADO ACTUAL
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando restauraci�n de DecimalSeparator"
    valorActualAntes = Application.DecimalSeparator
    caracterASCII = Asc(valorOriginal)
    
    fun801_LogMessage "[INICIO] " & strFuncion & " - Valor a restaurar: '" & valorOriginal & "' | ASCII: " & caracterASCII & _
                      " | Valor actual: '" & valorActualAntes & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 2: VALIDAR PAR�METRO DE ENTRADA
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Validando par�metro de entrada"
    fun801_LogMessage "[PASO 1] " & strContexto & " - Verificando formato y longitud (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) = 0 Then
        strMensajeError = "El valor original para DecimalSeparator no puede estar vac�o"
        Err.Raise ERROR_BASE_IMPORT + 1016, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[DETALLE] Longitud del valor: " & Len(valorOriginal) & " | Valor trimmed: '" & Trim(valorOriginal) & "'", False, "", strFuncion
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: VERIFICAR QUE SEA UN CAR�CTER V�LIDO
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Verificando validez del car�cter"
    fun801_LogMessage "[PASO 2] " & strContexto & " - Validando formato de separador decimal (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) <> 1 Then
        strMensajeError = "DecimalSeparator debe ser exactamente un car�cter - Longitud recibida: " & Len(valorOriginal) & " | Valor: '" & valorOriginal & "'"
        Err.Raise ERROR_BASE_IMPORT + 1017, strFuncion, strMensajeError
    End If
    
    If Not (valorOriginal = "." Or valorOriginal = ",") Then
        strMensajeError = "DecimalSeparator no es v�lido - Esperado: '.' o ',' | Recibido: '" & valorOriginal & "' | ASCII: " & caracterASCII
        Err.Raise ERROR_BASE_IMPORT + 1018, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[VALIDACI�N] Car�cter v�lido confirmado: '" & valorOriginal & "'", False, "", strFuncion
    
    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: APLICAR CONFIGURACI�N A EXCEL
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Aplicando nuevo separador decimal a Excel"
    fun801_LogMessage "[PASO 3] " & strContexto & " - Cambiando de '" & valorActualAntes & "' a '" & valorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    Application.DecimalSeparator = valorOriginal
    fun801_LogMessage "[APLICADO] Application.DecimalSeparator configurado", False, "", strFuncion
    
    lngLineaError = 160
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR QUE EL CAMBIO SE APLIC� CORRECTAMENTE
    '--------------------------------------------------------------------------
    lngLineaError = 170
    strContexto = "Verificando aplicaci�n correcta del separador"
    fun801_LogMessage "[PASO 4] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    valorActualDespues = Application.DecimalSeparator
    fun801_LogMessage "[VERIFICACI�N] Separador despu�s del cambio: '" & valorActualDespues & "' | Esperado: '" & valorOriginal & "'", False, "", strFuncion
    
    If valorActualDespues = valorOriginal Then
        fun801_LogMessage "[�XITO] DecimalSeparator restaurado exitosamente", False, "", strFuncion
        fun807_RestaurarDecimalSeparator = True
    Else
        strMensajeError = "Error en verificaci�n de DecimalSeparator - Valor esperado: '" & valorOriginal & "' | Valor actual: '" & valorActualDespues & "'"
        Err.Raise ERROR_BASE_IMPORT + 1019, strFuncion, strMensajeError
    End If
    
    lngLineaError = 180
    
    ' Verificaci�n adicional con formato de n�mero
    Dim numeroTest As Double
    Dim numeroFormateado As String
    numeroTest = 1.23
    numeroFormateado = CStr(numeroTest)
    fun801_LogMessage "[VERIFICACI�N ADICIONAL] N�mero test: " & numeroTest & " | Formateado: '" & numeroFormateado & "' | Separador detectado: '" & Mid(numeroFormateado, 2, 1) & "'", False, "", strFuncion
    
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Cambio aplicado: '" & valorActualAntes & "' ? '" & valorOriginal & "'", False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Valor Original: '" & valorOriginal & "'" & vbCrLf & _
                      "ASCII del valor: " & caracterASCII & vbCrLf & _
                      "Valor Antes: '" & valorActualAntes & "'" & vbCrLf & _
                      "Valor Despu�s: '" & valorActualDespues & "'" & vbCrLf & _
                      "Longitud valor: " & Len(valorOriginal) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun807_RestaurarDecimalSeparator = False
End Function

Public Function fun808_RestaurarThousandsSeparator(ByVal valorOriginal As String) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 808: RESTAURAR THOUSANDS SEPARATOR
    ' =============================================================================
    ' Fecha y hora de creaci�n: 2025-06-16 22:33:04 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Restaura la configuraci�n del separador de miles
    ' =============================================================================
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    Dim strContexto As String
    
    ' Variables de trabajo
    Dim valorActualAntes As String
    Dim valorActualDespues As String
    Dim caracterASCII As Integer
    Dim esCaracterValido As Boolean
    
    ' Inicializaci�n
    strFuncion = "fun808_RestaurarThousandsSeparator"
    fun808_RestaurarThousandsSeparator = False
    lngLineaError = 0
    strContexto = ""
    esCaracterValido = False
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' PASO 1: LOGGING INICIAL Y CAPTURA DE ESTADO ACTUAL
    '--------------------------------------------------------------------------
    lngLineaError = 100
    strContexto = "Iniciando restauraci�n de ThousandsSeparator"
    valorActualAntes = Application.ThousandsSeparator
    
    If Len(valorOriginal) > 0 Then
        caracterASCII = Asc(valorOriginal)
    Else
        caracterASCII = 0
    End If
    
    fun801_LogMessage "[INICIO] " & strFuncion & " - Valor a restaurar: '" & valorOriginal & "' | ASCII: " & caracterASCII & _
                      " | Valor actual: '" & valorActualAntes & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' PASO 2: VALIDAR PAR�METRO DE ENTRADA
    '--------------------------------------------------------------------------
    lngLineaError = 110
    strContexto = "Validando par�metro de entrada"
    fun801_LogMessage "[PASO 1] " & strContexto & " - Verificando formato y longitud (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) = 0 Then
        strMensajeError = "El valor original para ThousandsSeparator no puede estar vac�o"
        Err.Raise ERROR_BASE_IMPORT + 1020, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[DETALLE] Longitud del valor: " & Len(valorOriginal) & " | Valor trimmed: '" & Trim(valorOriginal) & "' | ASCII: " & caracterASCII, False, "", strFuncion
    
    lngLineaError = 120
    
    '--------------------------------------------------------------------------
    ' PASO 3: VERIFICAR QUE SEA UN CAR�CTER V�LIDO
    '--------------------------------------------------------------------------
    lngLineaError = 130
    strContexto = "Verificando validez del car�cter separador"
    fun801_LogMessage "[PASO 2] " & strContexto & " - Validando formato de separador de miles (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    If Len(Trim(valorOriginal)) <> 1 Then
        strMensajeError = "ThousandsSeparator debe ser exactamente un car�cter - Longitud recibida: " & Len(valorOriginal) & " | Valor: '" & valorOriginal & "'"
        Err.Raise ERROR_BASE_IMPORT + 1021, strFuncion, strMensajeError
    End If
    
    ' Verificar caracteres v�lidos para separador de miles
    If valorOriginal = "." Then
        esCaracterValido = True
        fun801_LogMessage "[VALIDACI�N] Car�cter punto (.) detectado", False, "", strFuncion
    ElseIf valorOriginal = "," Then
        esCaracterValido = True
        fun801_LogMessage "[VALIDACI�N] Car�cter coma (,) detectado", False, "", strFuncion
    ElseIf valorOriginal = " " Then
        esCaracterValido = True
        fun801_LogMessage "[VALIDACI�N] Car�cter espacio ( ) detectado | ASCII: 32", False, "", strFuncion
    ElseIf valorOriginal = Chr(39) Then ' Comilla simple
        esCaracterValido = True
        fun801_LogMessage "[VALIDACI�N] Car�cter comilla simple (') detectado | ASCII: 39", False, "", strFuncion
    Else
        esCaracterValido = False
        fun801_LogMessage "[VALIDACI�N] Car�cter no reconocido: '" & valorOriginal & "' | ASCII: " & caracterASCII, False, "", strFuncion
    End If
    
    If Not esCaracterValido Then
        strMensajeError = "ThousandsSeparator no es v�lido - Esperado: '.', ',', ' ' o ''' | Recibido: '" & valorOriginal & "' | ASCII: " & caracterASCII
        Err.Raise ERROR_BASE_IMPORT + 1022, strFuncion, strMensajeError
    End If
    
    fun801_LogMessage "[VALIDACI�N] Car�cter v�lido confirmado: '" & valorOriginal & "'", False, "", strFuncion
    
    lngLineaError = 140
    
    '--------------------------------------------------------------------------
    ' PASO 4: APLICAR CONFIGURACI�N A EXCEL
    '--------------------------------------------------------------------------
    lngLineaError = 150
    strContexto = "Aplicando nuevo separador de miles a Excel"
    fun801_LogMessage "[PASO 3] " & strContexto & " - Cambiando de '" & valorActualAntes & "' a '" & valorOriginal & "' (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    Application.ThousandsSeparator = valorOriginal
    fun801_LogMessage "[APLICADO] Application.ThousandsSeparator configurado", False, "", strFuncion
    
    lngLineaError = 160
    
    '--------------------------------------------------------------------------
    ' PASO 5: VERIFICAR QUE EL CAMBIO SE APLIC� CORRECTAMENTE
    '--------------------------------------------------------------------------
    lngLineaError = 170
    strContexto = "Verificando aplicaci�n correcta del separador"
    fun801_LogMessage "[PASO 4] " & strContexto & " (L�nea: " & lngLineaError & ")", False, "", strFuncion
    
    valorActualDespues = Application.ThousandsSeparator
    fun801_LogMessage "[VERIFICACI�N] Separador despu�s del cambio: '" & valorActualDespues & "' | ASCII: " & Asc(valorActualDespues) & " | Esperado: '" & valorOriginal & "'", False, "", strFuncion
    
    If valorActualDespues = valorOriginal Then
        fun801_LogMessage "[�XITO] ThousandsSeparator restaurado exitosamente", False, "", strFuncion
        fun808_RestaurarThousandsSeparator = True
    Else
        strMensajeError = "Error en verificaci�n de ThousandsSeparator - Valor esperado: '" & valorOriginal & "' (ASCII:" & caracterASCII & ") | Valor actual: '" & valorActualDespues & "' (ASCII:" & Asc(valorActualDespues) & ")"
        Err.Raise ERROR_BASE_IMPORT + 1023, strFuncion, strMensajeError
    End If
    
    lngLineaError = 180
    
    ' Verificaci�n adicional con formato de n�mero
    Dim numeroTest As Long
    Dim numeroFormateado As String
    numeroTest = 12345
    numeroFormateado = Format(numeroTest, "#,##0")
    fun801_LogMessage "[VERIFICACI�N ADICIONAL] N�mero test: " & numeroTest & " | Formateado: '" & numeroFormateado & "' | Longitud: " & Len(numeroFormateado), False, "", strFuncion
    
    If Len(numeroFormateado) >= 6 Then ' 12,345 = 6 caracteres m�nimo
        fun801_LogMessage "[VERIFICACI�N ADICIONAL] Separador en formato: '" & Mid(numeroFormateado, 3, 1) & "' | ASCII: " & Asc(Mid(numeroFormateado, 3, 1)), False, "", strFuncion
    End If
    
    fun801_LogMessage "[FINALIZACI�N] " & strFuncion & " completado exitosamente - Cambio aplicado: '" & valorActualAntes & "' ? '" & valorOriginal & "'", False, "", strFuncion
    
    Exit Function
    
GestorErrores:
    ' Construir mensaje de error exhaustivo
    strMensajeError = "[GESTOR DE ERRORES] Error en " & strFuncion & vbCrLf & _
                      "L�nea de Error: " & lngLineaError & vbCrLf & _
                      "Contexto: " & strContexto & vbCrLf & _
                      "N�mero de Error VBA: " & Err.Number & vbCrLf & _
                      "Descripci�n VBA: " & Err.Description & vbCrLf & _
                      "Valor Original: '" & valorOriginal & "'" & vbCrLf & _
                      "ASCII del valor: " & caracterASCII & vbCrLf & _
                      "Car�cter v�lido: " & esCaracterValido & vbCrLf & _
                      "Valor Antes: '" & valorActualAntes & "'" & vbCrLf & _
                      "Valor Despu�s: '" & valorActualDespues & "'" & vbCrLf & _
                      "Longitud valor: " & Len(valorOriginal) & vbCrLf & _
                      "Usuario: " & Environ("USERNAME") & vbCrLf & _
                      "Timestamp: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    fun801_LogMessage strMensajeError, True, "", strFuncion
    fun808_RestaurarThousandsSeparator = False
End Function

