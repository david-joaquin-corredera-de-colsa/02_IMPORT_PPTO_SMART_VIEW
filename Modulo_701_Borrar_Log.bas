Attribute VB_Name = "Modulo_701_Borrar_Log"
Option Explicit
Public Function Limpiar_Log() As Boolean
    
    '******************************************************************************
    ' FUNCI�N: Limpiar_Log
    ' FECHA Y HORA DE CREACI�N: 2025-06-14 14:09:13 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Funci�n para limpiar l�neas antiguas del log del sistema manteniendo un n�mero
    ' espec�fico de l�neas recientes. Implementa algoritmo avanzado de detecci�n de
    ' datos basado en conteo de filas vac�as consecutivas y eliminaci�n inteligente.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables de control y optimizaci�n de rendimiento
    ' 2. Configurar optimizaciones de rendimiento (pantalla, c�lculos, eventos)
    ' 3. Seleccionar y validar la hoja de log usando CONST_HOJA_LOG
    ' 4. Solicitar al usuario n�mero de l�neas a conservar con validaci�n exhaustiva
    ' 5. Inicializar contador de filas vac�as (vCounterFilasVacias)
    ' 6. Recorrer columna CONST_LOG_COLUMNA_FECHA_HORA para detectar �ltima fila real
    ' 7. Calcular primera l�nea de datos usando CONST_LOG_FILA_HEADERS+1
    ' 8. Determinar l�neas totales con datos excluyendo encabezados
    ' 9. Aplicar l�gica de eliminaci�n seg�n criterios establecidos
    ' 10. Eliminar rango de l�neas antiguas manteniendo las m�s recientes
    ' 11. Registrar operaci�n completada en el sistema de logging
    ' 12. Restaurar configuraciones de optimizaci�n originales
    ' 13. Manejo exhaustivo de errores con informaci�n detallada completa
    '
    ' PAR�METROS: Ninguno
    '
    ' VALOR DE RETORNO:
    ' - Boolean: True si la limpieza fue exitosa, False si error o cancelaci�n
    '
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    ' VERSI�N: 1.0
    '******************************************************************************
    
    ' Variables para control de errores
    Dim strFuncion As String
    Dim lngLineaError As Long
    Dim strMensajeError As String
    
    ' Variables para optimizaci�n de rendimiento
    Dim blnScreenUpdatingOriginal As Boolean
    Dim blnCalculationOriginal As Boolean
    Dim blnEventsOriginal As Boolean
    Dim blnDisplayAlertsOriginal As Boolean
    
    ' Variables principales del proceso
    Dim vLineasLogQueDejaremos As Long
    Dim vLineasLogTotales As Long
    Dim vPrimeraLineaLog As Long
    Dim vLineasAEliminar As Long
    Dim vUltimaFilaConDato As Long
    
    ' Variables para detecci�n avanzada de datos
    Dim vCounterFilasVacias As Integer
    Dim vFila As Long
    Dim vValorCelda As String
    
    ' Variables para manejo de hojas
    Dim wsLog As Worksheet
    Dim strNombreHojaLog As String
    Dim blnHojaLogExiste As Boolean
    
    ' Variables para interacci�n con usuario
    Dim strRespuestaUsuario As String
    Dim blnEntradaValida As Boolean
    Dim intIntentos As Integer
    
    ' Variables para eliminaci�n de filas
    Dim rngFilasAEliminar As Range
    Dim lngFilaInicio As Long
    Dim lngFilaFin As Long
    
    ' Inicializaci�n
    strFuncion = "Limpiar_Log"
    Limpiar_Log = False
    lngLineaError = 0
    intIntentos = 0
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables de control y optimizaci�n de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 50
    fun801_LogMessage "Iniciando limpieza de log del sistema", False, "", strFuncion
    
    ' Almacenar configuraciones originales para restaurar despu�s
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    blnDisplayAlertsOriginal = Application.DisplayAlerts
    
    '--------------------------------------------------------------------------
    ' 2. Configurar optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 60
    ' Desactivar actualizaci�n de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar c�lculo autom�tico para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    ' Desactivar alertas para eliminaci�n de filas
    Application.DisplayAlerts = False
    
    '--------------------------------------------------------------------------
    ' 3. Seleccionar y validar la hoja de log usando CONST_HOJA_LOG
    '--------------------------------------------------------------------------
    lngLineaError = 70
    ' Obtener nombre de hoja desde constante global
    strNombreHojaLog = CONST_HOJA_LOG
    
    ' Verificar existencia de la hoja de log
    blnHojaLogExiste = fun801_VerificarExistenciaHoja(ThisWorkbook, strNombreHojaLog)
    
    If Not blnHojaLogExiste Then
        Err.Raise ERROR_BASE_IMPORT + 1, strFuncion, _
            "La hoja de log no existe: " & Chr(34) & strNombreHojaLog & Chr(34)
    End If
    
    ' Obtener referencia a la hoja de log
    Set wsLog = ThisWorkbook.Worksheets(strNombreHojaLog)
    
    If wsLog Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 2, strFuncion, _
            "No se pudo obtener referencia a la hoja de log: " & Chr(34) & strNombreHojaLog & Chr(34)
    End If
    
    ' Seleccionar la hoja de log usando funci�n auxiliar existente
    If Not fun812_SeleccionarHoja(strNombreHojaLog) Then
        Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
            "No se pudo seleccionar la hoja de log: " & Chr(34) & strNombreHojaLog & Chr(34)
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Solicitar al usuario n�mero de l�neas a conservar con validaci�n exhaustiva
    '--------------------------------------------------------------------------
    lngLineaError = 80
    ' Reactivar temporalmente ScreenUpdating para mostrar InputBox
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    blnEntradaValida = False
    intIntentos = 0
    
    Do While Not blnEntradaValida And intIntentos < 5
        intIntentos = intIntentos + 1
        
        ' Solicitar entrada al usuario
        strRespuestaUsuario = InputBox( _
            "Cu�ntas l�neas antiguas desea dejar en el log?" & vbCrLf & vbCrLf & _
            "Introduzca un n�mero positivo:" & vbCrLf & _
            "(Intento " & intIntentos & " de 5)", _
            "Limpiar Log - " & strFuncion, _
            "100")
        
        ' Verificar si el usuario cancel�
        If Len(Trim(strRespuestaUsuario)) = 0 Then
            fun801_LogMessage "Usuario cancel� la operaci�n de limpieza de log", False, "", strFuncion
            GoTo RestaurarConfiguracion
        End If
        
        ' Validar entrada num�rica usando funci�n auxiliar mejorada
        If fun813_ValidarEntradaNumerica(strRespuestaUsuario, vLineasLogQueDejaremos) Then
            If vLineasLogQueDejaremos > 0 Then
                blnEntradaValida = True
                fun801_LogMessage "Usuario especific� conservar " & vLineasLogQueDejaremos & " l�neas", _
                    False, "", strFuncion
            Else
                MsgBox "El n�mero debe ser positivo (mayor que 0)." & vbCrLf & _
                       "Por favor, intente nuevamente.", vbExclamation, "Entrada Inv�lida"
            End If
        Else
            MsgBox "Entrada inv�lida. Debe introducir un n�mero entero positivo." & vbCrLf & _
                   "Por favor, intente nuevamente.", vbExclamation, "Entrada Inv�lida"
        End If
    Loop
    
    ' Verificar si se agotaron los intentos
    If Not blnEntradaValida Then
        fun801_LogMessage "Se agotaron los intentos de entrada del usuario", True, "", strFuncion
        GoTo RestaurarConfiguracion
    End If
    
    ' Desactivar nuevamente ScreenUpdating
    Application.ScreenUpdating = False
    
    '--------------------------------------------------------------------------
    ' 5. Inicializar contador de filas vac�as (vCounterFilasVacias)
    '--------------------------------------------------------------------------
    lngLineaError = 90
    ' Crear variable vCounterFilasVacias tipo Integer y inicializar con valor 0
    vCounterFilasVacias = 0
    vUltimaFilaConDato = 0
    vFila = 1
    
    fun801_LogMessage "Inicializando detecci�n de datos: vCounterFilasVacias = " & vCounterFilasVacias, _
        False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 6. Recorrer columna CONST_LOG_COLUMNA_FECHA_HORA para detectar �ltima fila real
    '--------------------------------------------------------------------------
    lngLineaError = 100
    fun801_LogMessage "Iniciando recorrido de columna " & CONST_LOG_COLUMNA_FECHA_HORA & _
        " para detectar �ltima fila con datos", False, "", strFuncion
    
    ' Recorrer desde vFila = 1 hasta que vCounterFilasVacias = 10
    Do While vCounterFilasVacias < 10 And vFila <= 1048576  ' L�mite m�ximo Excel
        
        '----------------------------------------------------------------------
        ' 6.1. Enviar valor de celda a variable vValorCelda tipo String
        '----------------------------------------------------------------------
        vValorCelda = CStr(wsLog.Cells(vFila, CONST_LOG_COLUMNA_FECHA_HORA).Value)
        
        '----------------------------------------------------------------------
        ' 6.2. Hacer vValorCelda = Trim(vValorCelda)
        '----------------------------------------------------------------------
        vValorCelda = Trim(vValorCelda)
        
        '----------------------------------------------------------------------
        ' 6.3. Si vValorCelda<>"" entonces actualizar vUltimaFilaConDato y resetear contador
        '----------------------------------------------------------------------
        If vValorCelda <> "" Then
            vUltimaFilaConDato = vFila
            vCounterFilasVacias = 0
        Else
            '----------------------------------------------------------------------
            ' 6.4. Si vValorCelda="" entonces incrementar contador y ajustar vUltimaFilaConDato
            '----------------------------------------------------------------------
            vCounterFilasVacias = vCounterFilasVacias + 1
            vUltimaFilaConDato = vFila - vCounterFilasVacias
        End If
        
        ' Incrementar fila para siguiente iteraci�n
        vFila = vFila + 1
        
        ' Log peri�dico cada 1000 filas para debugging en logs grandes
        If vFila Mod 1000 = 0 Then
            fun801_LogMessage "Procesando fila " & vFila & " - Filas vac�as consecutivas: " & _
                vCounterFilasVacias & " - �ltima con datos: " & vUltimaFilaConDato, False, "", strFuncion
        End If
    Loop
    
    'MsgBox "vUltimaFilaConDato=" & vUltimaFilaConDato
    
    fun801_LogMessage "Detecci�n completada - �ltima fila con datos: " & vUltimaFilaConDato & _
        " - Filas vac�as consecutivas al final: " & vCounterFilasVacias & _
        " - Total filas examinadas: " & (vFila - 1), False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 7. Calcular primera l�nea de datos usando CONST_LOG_FILA_HEADERS+1
    '--------------------------------------------------------------------------
    lngLineaError = 110
    ' Almacenar valor CONST_LOG_FILA_HEADERS+1 en variable vPrimeraLineaLog
    vPrimeraLineaLog = CONST_LOG_FILA_HEADERS + 1
    
    fun801_LogMessage "Primera l�nea de datos calculada: " & vPrimeraLineaLog & _
        " (CONST_LOG_FILA_HEADERS: " & CONST_LOG_FILA_HEADERS & ")", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 8. Determinar l�neas totales con datos excluyendo encabezados
    '--------------------------------------------------------------------------
    lngLineaError = 120
    ' Calcular l�neas totales de datos (excluyendo encabezados)
    If vUltimaFilaConDato >= vPrimeraLineaLog Then
        vLineasLogTotales = vUltimaFilaConDato - vPrimeraLineaLog + 1
    Else
        vLineasLogTotales = 0
    End If
    
    fun801_LogMessage "L�neas totales con datos calculadas: " & vLineasLogTotales & _
        " (�ltima fila datos: " & vUltimaFilaConDato & ", Primera l�nea datos: " & vPrimeraLineaLog & ")", _
        False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 9. Aplicar l�gica de eliminaci�n seg�n criterios establecidos
    '--------------------------------------------------------------------------
    lngLineaError = 130
    
    ' Aplicar l�gica seg�n especificaciones: 0.5.6 y 0.5.7
    If vLineasLogTotales < vLineasLogQueDejaremos Then
        ' Especificaci�n 0.5.6: Si vLineasLogTotales<vLineasLogQueDejaremos entonces no borrar ninguna l�nea
        vLineasAEliminar = 0
        fun801_LogMessage "No es necesario eliminar l�neas. Total: " & vLineasLogTotales & _
            " < A conservar: " & vLineasLogQueDejaremos, False, "", strFuncion
        
    ElseIf vLineasLogTotales > vLineasLogQueDejaremos Then
        ' Especificaci�n 0.5.7: Si vLineasLogTotales>vLineasLogQueDejaremos entonces borrar l�neas
        ' Desde vPrimeraLineaLog hasta dejar solo vLineasLogQueDejaremos
        vLineasAEliminar = vLineasLogTotales - vLineasLogQueDejaremos
        
        ' Calcular rango de filas a eliminar (desde las m�s antiguas)
        lngFilaInicio = vPrimeraLineaLog
        lngFilaFin = vPrimeraLineaLog + vLineasAEliminar - 1
        
        fun801_LogMessage "Se eliminar�n " & vLineasAEliminar & " l�neas antiguas (filas " & _
            lngFilaInicio & " a " & lngFilaFin & ") para dejar exactamente " & vLineasLogQueDejaremos & " l�neas", _
            False, "", strFuncion
    Else
        ' Caso especial: vLineasLogTotales = vLineasLogQueDejaremos (no eliminar nada)
        vLineasAEliminar = 0
        fun801_LogMessage "Las l�neas actuales (" & vLineasLogTotales & _
            ") coinciden exactamente con las solicitadas (" & vLineasLogQueDejaremos & "). No se eliminar�n l�neas.", _
            False, "", strFuncion
    End If
    
    '--------------------------------------------------------------------------
    ' 10. Eliminar rango de l�neas antiguas manteniendo las m�s recientes
    '--------------------------------------------------------------------------
    lngLineaError = 140
    
    If vLineasAEliminar > 0 Then
        fun801_LogMessage "Procediendo a eliminar " & vLineasAEliminar & " l�neas antiguas", _
            False, "", strFuncion
        
        ' Crear referencia al rango de filas a eliminar
        Set rngFilasAEliminar = wsLog.Rows(lngFilaInicio & ":" & lngFilaFin)
        
        ' Eliminar las filas (m�todo compatible con Excel 97-365)
        rngFilasAEliminar.Delete Shift:=xlUp
        
        fun801_LogMessage "L�neas eliminadas exitosamente. Filas " & lngFilaInicio & _
            " a " & lngFilaFin & " han sido removidas. Quedan exactamente " & vLineasLogQueDejaremos & " l�neas", _
            False, "", strFuncion
    End If
    
    '--------------------------------------------------------------------------
    ' 11. Registrar operaci�n completada en el sistema de logging
    '--------------------------------------------------------------------------
    lngLineaError = 150
    fun801_LogMessage "Limpieza de log completada exitosamente. " & _
        "L�neas conservadas: " & vLineasLogQueDejaremos & _
        ", L�neas eliminadas: " & vLineasAEliminar & _
        ", Total original: " & vLineasLogTotales & _
        ", M�todo detecci�n: contador filas vac�as", False, "", strFuncion
    
    '--------------------------------------------------------------------------
    ' 12. Proceso completado exitosamente - Especificaci�n 0.5.6
    '--------------------------------------------------------------------------
    lngLineaError = 160
    Limpiar_Log = True
    
RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 13. Restaurar configuraciones de optimizaci�n originales
    '--------------------------------------------------------------------------
    lngLineaError = 170
    ' Restaurar configuraci�n original de alertas
    Application.DisplayAlerts = blnDisplayAlertsOriginal
    
    ' Restaurar configuraci�n original de eventos
    Application.EnableEvents = blnEventsOriginal
    
    ' Restaurar configuraci�n original de c�lculo
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
    ' Restaurar configuraci�n original de actualizaci�n de pantalla
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Limpiar referencias de objetos
    Set wsLog = Nothing
    Set rngFilasAEliminar = Nothing
    
    'Retornamos a la hoja principal
    ThisWorkbook.Worksheets(CONST_HOJA_EJECUTAR_PROCESOS).Select
    ThisWorkbook.Worksheets(CONST_HOJA_EJECUTAR_PROCESOS).Activate
    
    fun801_LogMessage "Funci�n Limpiar_Log finalizada con resultado: " & Limpiar_Log, _
        False, "", strFuncion
                
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' 14. Manejo exhaustivo de errores con informaci�n detallada completa
    '--------------------------------------------------------------------------
    
    ' Construir mensaje de error detallado - Especificaci�n 0.5.7
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja de log: " & Chr(34) & strNombreHojaLog & Chr(34) & vbCrLf & _
                      "L�neas a conservar: " & vLineasLogQueDejaremos & vbCrLf & _
                      "L�neas totales encontradas: " & vLineasLogTotales & vbCrLf & _
                      "Primera l�nea de datos: " & vPrimeraLineaLog & vbCrLf & _
                      "�ltima fila con dato: " & vUltimaFilaConDato & vbCrLf & _
                      "Contador filas vac�as: " & vCounterFilasVacias & vbCrLf & _
                      "Fila actual procesamiento: " & vFila & vbCrLf & _
                      "Intentos de entrada: " & intIntentos & vbCrLf & _
                      "Fecha y Hora: " & Now()
    
    ' Registrar error en log del sistema
    fun801_LogMessage strMensajeError, True, "", strFuncion
    
    ' Log del error para debugging
    Debug.Print strMensajeError
    
    ' Restaurar configuraciones en caso de error
    On Error Resume Next
    Application.DisplayAlerts = blnDisplayAlertsOriginal
    Application.EnableEvents = blnEventsOriginal
    If blnCalculationOriginal Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    Application.ScreenUpdating = blnScreenUpdatingOriginal
    
    ' Limpiar referencias de objetos
    Set wsLog = Nothing
    Set rngFilasAEliminar = Nothing
    
    ' Retornar False para indicar error - Especificaci�n 0.5.7
    Limpiar_Log = False
    
    ' Ir a restaurar configuraci�n si a�n no se hizo
    If lngLineaError < 170 Then
        Resume RestaurarConfiguracion
    End If
    
    'Retornamos a la hoja principal
    ThisWorkbook.Worksheets(CONST_HOJA_EJECUTAR_PROCESOS).Select
    ThisWorkbook.Worksheets(CONST_HOJA_EJECUTAR_PROCESOS).Activate
    
End Function
Public Function fun812_SeleccionarHoja(ByVal strNombreHoja As String) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun812_SeleccionarHoja
    ' FECHA Y HORA DE CREACI�N: 2025-06-14 14:09:13 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Selecciona una hoja espec�fica en ThisWorkbook de forma segura
    ' Reutiliza la l�gica de validaci�n del proyecto existente
    '
    ' PAR�METROS:
    ' - strNombreHoja (String): Nombre de la hoja a seleccionar
    '
    ' RETORNA: Boolean - True si �xito, False si error
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim strFuncion As String
    
    ' Inicializaci�n
    strFuncion = "fun812_SeleccionarHoja"
    fun812_SeleccionarHoja = False
    
    ' Validar par�metro de entrada
    If Len(Trim(strNombreHoja)) = 0 Then
        fun801_LogMessage "Par�metro strNombreHoja est� vac�o", True, "", strFuncion
        Exit Function
    End If
    
    ' Verificar existencia usando funci�n existente del proyecto
    If Not fun801_VerificarExistenciaHoja(ThisWorkbook, strNombreHoja) Then
        fun801_LogMessage "La hoja no existe: " & Chr(34) & strNombreHoja & Chr(34), True, "", strFuncion
        Exit Function
    End If
    
    ' Obtener referencia y seleccionar
    Set ws = ThisWorkbook.Worksheets(strNombreHoja)
    ws.Activate
    ws.Range("A1").Select
    
    fun812_SeleccionarHoja = True
    Exit Function
    
ErrorHandler:
    fun801_LogMessage "Error en " & strFuncion & ": " & Err.Description, True, "", strFuncion
    fun812_SeleccionarHoja = False
End Function
Public Function fun813_ValidarEntradaNumerica(ByVal strEntrada As String, ByRef lngResultado As Long) As Boolean
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun813_ValidarEntradaNumerica
    ' FECHA Y HORA DE CREACI�N: 2025-06-14 14:09:13 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Valida que una cadena sea un n�mero entero v�lido y positivo
    ' Compatible con Excel 97-365, validaciones exhaustivas
    '
    ' PAR�METROS:
    ' - strEntrada (String): Cadena a validar
    ' - lngResultado (Long): Variable donde almacenar resultado por referencia
    '
    ' RETORNA: Boolean - True si es n�mero v�lido, False si no
    ' COMPATIBILIDAD: Excel 97, 2003, 2007, 365, OneDrive, SharePoint, Teams
    '******************************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim strLimpia As String
    Dim i As Integer
    Dim strCaracter As String
    
    ' Inicializaci�n
    fun813_ValidarEntradaNumerica = False
    lngResultado = 0
    
    ' Limpiar entrada
    strLimpia = Trim(strEntrada)
    
    ' Validar no vac�a
    If Len(strLimpia) = 0 Then Exit Function
    
    ' Validar longitud razonable
    If Len(strLimpia) > 9 Then Exit Function
    
    ' Validar caracteres num�ricos
    For i = 1 To Len(strLimpia)
        strCaracter = Mid(strLimpia, i, 1)
        If strCaracter < "0" Or strCaracter > "9" Then Exit Function
    Next i
    
    ' Convertir a n�mero
    On Error Resume Next
    lngResultado = CLng(strLimpia)
    If Err.Number <> 0 Then
        lngResultado = 0
        On Error GoTo ErrorHandler
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Validar rango razonable
    If lngResultado < 0 Or lngResultado > 1000000 Then
        lngResultado = 0
        Exit Function
    End If
    
    fun813_ValidarEntradaNumerica = True
    Exit Function
    
ErrorHandler:
    fun813_ValidarEntradaNumerica = False
    lngResultado = 0
End Function

Public Function Limpiar_Otra_Informacion() As Boolean
    
    '******************************************************************************
    ' FUNCI�N: Limpiar_Otra_Informacion
    ' FECHA Y HORA DE CREACI�N: 2025-06-15 09:40:56 UTC
    ' AUTOR: david-joaquin-corredera-de-colsa
    '
    ' DESCRIPCI�N:
    ' Limpia de forma segura el contenido de la celda especificada por CONST_CELDA_USERNAME
    ' en la hoja designada por CONST_HOJA_USERNAME, preservando el estado de visibilidad
    ' original de la hoja. Funci�n auxiliar para operaciones de limpieza de datos
    ' sensibles del sistema.
    '
    ' RESUMEN EXHAUSTIVO DE PASOS:
    ' 1. Inicializaci�n de variables de control de errores y optimizaci�n
    ' 2. Configuraci�n de optimizaciones de rendimiento (pantalla, c�lculos)
    ' 3. Verificaci�n de existencia de la hoja objetivo usando constantes del sistema
    ' 4. Detecci�n y almacenamiento del estado de visibilidad actual de la hoja
    ' 5. Hacer temporalmente visible la hoja si est� oculta para permitir acceso
    ' 6. Obtenci�n de referencia segura a la hoja de trabajo
    ' 7. Validaci�n y limpieza del contenido del rango de celdas especificado
    ' 8. Restauraci�n del estado de visibilidad original de la hoja
    ' 9. Restauraci�n de configuraciones de optimizaci�n del sistema
    ' 10. Logging de operaci�n y manejo exhaustivo de errores
    '
    ' PAR�METROS: Ninguno (usa constantes globales del sistema)
    ' RETORNA: Boolean - True si la operaci�n se complet� exitosamente, False si error
    '
    ' CONSTANTES UTILIZADAS:
    ' - CONST_HOJA_USERNAME: Nombre de la hoja que contiene la informaci�n
    ' - CONST_CELDA_USERNAME: Direcci�n de la celda a limpiar
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
    
    ' Variables para manejo de hojas y visibilidad
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim blnHojaExiste As Boolean
    Dim intVisibilidadOriginal As Integer
    Dim blnCambioVisibilidad As Boolean
    Dim rngCelda As Range
    
    ' Inicializaci�n
    strFuncion = "Limpiar_Otra_Informacion"
    Limpiar_Otra_Informacion = False
    lngLineaError = 0
    blnCambioVisibilidad = False
    intVisibilidadOriginal = xlSheetVisible
    
    On Error GoTo GestorErrores
    
    '--------------------------------------------------------------------------
    ' 1. Inicializaci�n de variables de control de errores y optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 30
    
    ' Registrar inicio de operaci�n
    Call fun801_LogMessage("Iniciando limpieza de informaci�n en hoja: " & _
        CONST_HOJA_USERNAME & ", celda: " & CONST_CELDA_USERNAME, False, "", strFuncion)
    
    ' Almacenar configuraciones originales para restaurar despu�s
    blnScreenUpdatingOriginal = Application.ScreenUpdating
    blnCalculationOriginal = (Application.Calculation = xlCalculationAutomatic)
    blnEventsOriginal = Application.EnableEvents
    
    '--------------------------------------------------------------------------
    ' 2. Configuraci�n de optimizaciones de rendimiento
    '--------------------------------------------------------------------------
    lngLineaError = 40
    
    ' Desactivar actualizaci�n de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    
    ' Desactivar c�lculo autom�tico para mayor velocidad
    Application.Calculation = xlCalculationManual
    
    ' Desactivar eventos para evitar interferencias
    Application.EnableEvents = False
    
    '--------------------------------------------------------------------------
    ' 3. Verificaci�n de existencia de la hoja objetivo
    '--------------------------------------------------------------------------
    lngLineaError = 50
    
    ' Obtener referencia al libro actual
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
       
    ' Verificar que tenemos una referencia v�lida al libro
    If wb Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 9001, strFuncion, _
            "No se pudo obtener referencia al libro de trabajo"
    End If
    
    ' Verificar existencia de la hoja usando funci�n auxiliar existente del proyecto
    blnHojaExiste = fun801_VerificarExistenciaHoja(wb, CONST_HOJA_USERNAME)
    
    If Not blnHojaExiste Then
        Err.Raise ERROR_BASE_IMPORT + 9002, strFuncion, _
            "La hoja especificada no existe: " & CONST_HOJA_USERNAME
    End If
    
    '--------------------------------------------------------------------------
    ' 4. Detecci�n y almacenamiento del estado de visibilidad actual
    '--------------------------------------------------------------------------
    lngLineaError = 60
    
    ' Obtener referencia a la hoja sin cambiar su visibilidad a�n
    Set ws = wb.Worksheets(CONST_HOJA_USERNAME)
    
    If ws Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 9003, strFuncion, _
            "No se pudo obtener referencia a la hoja: " & CONST_HOJA_USERNAME
    End If
    
    ' Almacenar estado de visibilidad original
    intVisibilidadOriginal = ws.Visible
    
    Call fun801_LogMessage("Estado de visibilidad original detectado: " & _
        CStr(intVisibilidadOriginal) & " para hoja: " & CONST_HOJA_USERNAME, _
        False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 5. Hacer temporalmente visible la hoja si est� oculta
    '--------------------------------------------------------------------------
    lngLineaError = 70
    
    ' Verificar si la hoja est� oculta y necesita hacerse visible temporalmente
    If ws.Visible <> xlSheetVisible Then
        ' Verificar que el libro permite cambiar visibilidad
        If Not wb.ProtectStructure Then
            ws.Visible = xlSheetVisible
            blnCambioVisibilidad = True
            Call fun801_LogMessage("Hoja hecha temporalmente visible para acceso: " & _
                CONST_HOJA_USERNAME, False, "", strFuncion)
        Else
            Call fun801_LogMessage("ADVERTENCIA - Libro protegido, no se puede cambiar visibilidad", _
                False, "", strFuncion)
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' 6. Obtenci�n de referencia segura a la hoja de trabajo
    '--------------------------------------------------------------------------
    lngLineaError = 80
    
    ' Verificar que la hoja es accesible
    If ws.ProtectContents Then
        Call fun801_LogMessage("ADVERTENCIA - Hoja protegida contra escritura: " & _
            CONST_HOJA_USERNAME, False, "", strFuncion)
    End If
    
    '--------------------------------------------------------------------------
    ' 7. Validaci�n y limpieza del contenido del rango de celdas
    '--------------------------------------------------------------------------
    lngLineaError = 90
    
    ' Obtener referencia al rango de celdas especificado
    Set rngCelda = ws.Range(CONST_CELDA_USERNAME)
    
    If rngCelda Is Nothing Then
        Err.Raise ERROR_BASE_IMPORT + 9004, strFuncion, _
            "No se pudo obtener referencia al rango: " & CONST_CELDA_USERNAME
    End If
    
    ' Limpiar el contenido del rango (no el formato)
    rngCelda.ClearContents
    
    Call fun801_LogMessage("Contenido limpiado exitosamente en rango: " & _
        CONST_CELDA_USERNAME & " de hoja: " & CONST_HOJA_USERNAME, False, "", strFuncion)
    
    '--------------------------------------------------------------------------
    ' 8. Restauraci�n del estado de visibilidad original
    '--------------------------------------------------------------------------
    lngLineaError = 100
    
    ' Restaurar visibilidad original si se cambi�
    If blnCambioVisibilidad And Not wb.ProtectStructure Then
        ws.Visible = intVisibilidadOriginal
        Call fun801_LogMessage("Estado de visibilidad restaurado a: " & _
            CStr(intVisibilidadOriginal) & " para hoja: " & CONST_HOJA_USERNAME, _
            False, "", strFuncion)
    End If
    
    '--------------------------------------------------------------------------
    ' 9. Proceso completado exitosamente
    '--------------------------------------------------------------------------
    lngLineaError = 110
    Limpiar_Otra_Informacion = True
    
RestaurarConfiguracion:
    '--------------------------------------------------------------------------
    ' 10. Restauraci�n de configuraciones de optimizaci�n
    '--------------------------------------------------------------------------
    lngLineaError = 120
    
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
    Set rngCelda = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    Call fun801_LogMessage("Operaci�n de limpieza completada exitosamente", _
        False, "", strFuncion)
    Exit Function

GestorErrores:
    '--------------------------------------------------------------------------
    ' Manejo exhaustivo de errores con informaci�n detallada
    '--------------------------------------------------------------------------
    
    ' Intentar restaurar visibilidad en caso de error
    On Error Resume Next
    If blnCambioVisibilidad And Not ws Is Nothing And Not wb Is Nothing Then
        If Not wb.ProtectStructure Then
            ws.Visible = intVisibilidadOriginal
        End If
    End If
    On Error GoTo 0
    
    ' Construir mensaje de error detallado
    strMensajeError = "Error en " & strFuncion & vbCrLf & _
                      "L�nea: " & lngLineaError & vbCrLf & _
                      "N�mero de Error: " & Err.Number & vbCrLf & _
                      "Descripci�n: " & Err.Description & vbCrLf & _
                      "Hoja objetivo: " & CONST_HOJA_USERNAME & vbCrLf & _
                      "Celda objetivo: " & CONST_CELDA_USERNAME & vbCrLf & _
                      "Visibilidad original: " & CStr(intVisibilidadOriginal) & vbCrLf & _
                      "Cambio visibilidad: " & blnCambioVisibilidad & vbCrLf & _
                      "Fecha y Hora: " & Now() & vbCrLf & _
                      "Compatibilidad: Excel 97/2003/2007/365, OneDrive/SharePoint/Teams"
    
    ' Registrar error en log del sistema
    Call fun801_LogMessage(strMensajeError, True, "", strFuncion)
    
    ' Para debugging en desarrollo
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
    Set rngCelda = Nothing
    Set ws = Nothing
    Set wb = Nothing
    
    ' Retornar False para indicar error
    Limpiar_Otra_Informacion = False
End Function

