Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_02"

Option Explicit

Public Function fun802_SheetExists(ByVal strSheetName As String) As Boolean
    
    '========================================================================
    ' FUNCION AUXILIAR: fun802_SheetExists
    ' Descripcion : Verifica de forma segura si existe una hoja (worksheet)
    '               con el nombre indicado en el libro actual
    '               antes de entrar a trabajar con ella
    ' Fecha       : 2025-06-01
    ' Retorna     : Boolean
    '========================================================================
    
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    fun802_SheetExists = False
    Set ws = ThisWorkbook.Worksheets(strSheetName)
    If Not ws Is Nothing Then
        fun802_SheetExists = True
    End If
    Exit Function
ErrorHandler:
    fun802_SheetExists = False
End Function

Public Function fun811_DetectarThousandsSeparatorLegacy() As String

    ' =============================================================================
    ' FUNCI�N AUXILIAR 811: DETECTAR THOUSANDS SEPARATOR (M�TODO LEGACY)
    ' =============================================================================
    ' Fecha: 2025-05-26 17:43:59 UTC
    ' Descripci�n: M�todo alternativo para detectar separador de miles en versiones antiguas
    ' Par�metros: Ninguno
    ' Retorna: String (car�cter del separador de miles)
    ' Compatibilidad: Excel 97, 2003
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    ' Variables para detecci�n
    Dim numeroFormateado As String
    Dim lineaError As Long
    
    lineaError = 1200
    
    ' M�todo alternativo: formatear un n�mero grande y extraer el separador
    ' Compatible con Excel 97 y versiones antiguas
    numeroFormateado = Format(1000, "#,##0")
    
    lineaError = 1210
    
    ' El separador de miles es el segundo car�cter en n�meros de 4 d�gitos
    If Len(numeroFormateado) >= 2 Then
        fun811_DetectarThousandsSeparatorLegacy = Mid(numeroFormateado, 2, 1)
    Else
        ' Si no hay separador visible, asumir coma por defecto
        fun811_DetectarThousandsSeparatorLegacy = ","
    End If
    
    lineaError = 1220
    
    Exit Function
    
ErrorHandler:
    ' En caso de error, asumir coma por defecto
    fun811_DetectarThousandsSeparatorLegacy = ","
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: fun811_DetectarThousandsSeparatorLegacy" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function





















Public Function borrame_fun809_OcultarHojaDelimitadores(ws As Worksheet) As Boolean
    
    ' =============================================================================
    ' FUNCI�N AUXILIAR 809: OCULTAR HOJA DE DELIMITADORES
    ' =============================================================================
    ' Fecha: 2025-05-26 18:41:20 UTC
    ' Usuario: david-joaquin-corredera-de-colsa
    ' Descripci�n: Oculta la hoja de delimitadores si est� habilitada la opci�n
    ' Par�metros: ws (Worksheet)
    ' Retorna: Boolean (True si exitoso, False si error)
    ' Compatibilidad: Excel 97, 2003, 365, OneDrive, SharePoint, Teams
    ' =============================================================================
    
    On Error GoTo ErrorHandler
    
    Dim lineaError As Long
    
    lineaError = 1000
    borrame_fun809_OcultarHojaDelimitadores = True
    
    ' Verificar par�metro de entrada
    If ws Is Nothing Then
        borrame_fun809_OcultarHojaDelimitadores = False
        Exit Function
    End If
    
    lineaError = 1010
    
    ' Verificar que el libro permite ocultar hojas (no protegido)
    If ws.Parent.ProtectStructure Then
        Debug.Print "ADVERTENCIA: No se puede ocultar hoja, libro protegido - Funci�n: borrame_fun809_OcultarHojaDelimitadores - " & Now()
        Exit Function
    End If
    
    lineaError = 1020
    
    ' Ocultar la hoja (compatible con todas las versiones de Excel)
    ws.Visible = xlSheetHidden
    Debug.Print "INFO: Hoja " & ws.Name & " ocultada - Funci�n: borrame_fun809_OcultarHojaDelimitadores - " & Now()
    
    lineaError = 1030
    
    Exit Function
    
ErrorHandler:
    borrame_fun809_OcultarHojaDelimitadores = False
    
    ' Informaci�n detallada del error
    Dim mensajeError As String
    mensajeError = "ERROR EN FUNCI�N: borrame_fun809_OcultarHojaDelimitadores" & vbCrLf & _
                   "TIPO DE ERROR: " & Err.Number & " - " & Err.Description & vbCrLf & _
                   "L�NEA DE ERROR APROXIMADA: " & lineaError & vbCrLf & _
                   "L�NEA VBA: " & Erl & vbCrLf & _
                   "HOJA: " & ws.Name & vbCrLf & _
                   "FECHA Y HORA: " & Now()
    
    Debug.Print mensajeError
    
End Function


Public Function fun802_VerificarCompatibilidad() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun802_VerificarCompatibilidad
    ' PROP�SITO: Verifica compatibilidad con diferentes versiones de Excel
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = compatible, False = no compatible)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun802
    
    Dim strVersionExcel As String
    Dim dblVersionNumero As Double
    
    ' Obtener versi�n de Excel
    strVersionExcel = Application.Version
    dblVersionNumero = CDbl(strVersionExcel)
    
    ' Verificar compatibilidad (Excel 97 = 8.0, 2003 = 11.0, 365 = 16.0+)
    If dblVersionNumero >= 8# Then
        fun802_VerificarCompatibilidad = True
    Else
        fun802_VerificarCompatibilidad = False
    End If
    
    Exit Function

ErrorHandler_fun802:
    ' En caso de error, asumir compatibilidad
    fun802_VerificarCompatibilidad = True
End Function

Public Sub fun803_ObtenerConfiguracionActual(ByRef strDecimalAnterior As String, ByRef strMilesAnterior As String)
    ' =============================================================================
    ' FUNCI�N: fun803_ObtenerConfiguracionActual
    ' PROP�SITO: Obtiene la configuraci�n actual de delimitadores
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error GoTo ErrorHandler_fun803
    
    ' Obtener delimitador decimal actual
    strDecimalAnterior = Application.International(xlDecimalSeparator)
    
    ' Obtener delimitador de miles actual
    strMilesAnterior = Application.International(xlThousandsSeparator)
    
    Exit Sub

ErrorHandler_fun803:
    ' En caso de error, usar valores por defecto
    strDecimalAnterior = "."
    strMilesAnterior = ","
End Sub

Public Function fun804_AplicarNuevosDelimitadores() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun804_AplicarNuevosDelimitadores
    ' PROP�SITO: Aplica los nuevos delimitadores al sistema
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = �xito, False = error)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun804
    
    ' Aplicar nuevo delimitador decimal
    Application.DecimalSeparator = vDelimitadorDecimal_HFM
    
    ' Aplicar nuevo delimitador de miles
    Application.ThousandsSeparator = vDelimitadorMiles_HFM
    
    ' Forzar que Excel use los delimitadores del sistema
    Application.UseSystemSeparators = False
    
    ' Actualizar la pantalla
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    fun804_AplicarNuevosDelimitadores = True
    Exit Function

ErrorHandler_fun804:
    fun804_AplicarNuevosDelimitadores = False
End Function

Public Function fun805_VerificarAplicacionDelimitadores() As Boolean
    ' =============================================================================
    ' FUNCI�N: fun805_VerificarAplicacionDelimitadores
    ' PROP�SITO: Verifica que los delimitadores se aplicaron correctamente
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' RETORNA: Boolean (True = aplicados correctamente, False = error)
    ' =============================================================================
    On Error GoTo ErrorHandler_fun805
    
    Dim strDecimalActual As String
    Dim strMilesActual As String
    
    ' Obtener delimitadores actuales
    strDecimalActual = Application.DecimalSeparator
    strMilesActual = Application.ThousandsSeparator
    
    ' Verificar que coinciden con los deseados
    If strDecimalActual = vDelimitadorDecimal_HFM And strMilesActual = vDelimitadorMiles_HFM Then
        fun805_VerificarAplicacionDelimitadores = True
    Else
        fun805_VerificarAplicacionDelimitadores = False
    End If
    
    Exit Function

ErrorHandler_fun805:
    fun805_VerificarAplicacionDelimitadores = False
End Function

Public Sub fun806_RestaurarConfiguracion(ByVal strDecimalAnterior As String, ByVal strMilesAnterior As String)
    ' =============================================================================
    ' FUNCI�N: fun806_RestaurarConfiguracion
    ' PROP�SITO: Restaura la configuraci�n anterior en caso de error
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    On Error Resume Next
    
    ' Restaurar delimitador decimal anterior
    Application.DecimalSeparator = strDecimalAnterior
    
    ' Restaurar delimitador de miles anterior
    Application.ThousandsSeparator = strMilesAnterior
    
    ' Restaurar uso de separadores del sistema
    Application.UseSystemSeparators = True
    
    On Error GoTo 0
End Sub

Public Sub fun807_MostrarErrorDetallado(ByVal strFuncion As String, ByVal strTipoError As String, _
                                        ByVal lngLinea As Long, ByVal lngNumeroError As Long, _
                                        ByVal strDescripcionError As String)
    
    ' =============================================================================
    ' FUNCI�N: fun807_MostrarErrorDetallado
    ' PROP�SITO: Muestra informaci�n detallada del error ocurrido
    ' FECHA: 2025-05-26 15:11:21 UTC
    ' =============================================================================
    Dim strMensajeError As String
    
    ' Construir mensaje de error detallado
    strMensajeError = "ERROR EN FUNCI�N DE DELIMITADORES" & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Funci�n: " & strFuncion & vbCrLf
    strMensajeError = strMensajeError & "Tipo de Error: " & strTipoError & vbCrLf
    strMensajeError = strMensajeError & "L�nea Aproximada: " & CStr(lngLinea) & vbCrLf
    strMensajeError = strMensajeError & "N�mero de Error VBA: " & CStr(lngNumeroError) & vbCrLf
    strMensajeError = strMensajeError & "Descripci�n: " & strDescripcionError & vbCrLf & vbCrLf
    strMensajeError = strMensajeError & "Fecha/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    ' Mostrar mensaje de error
    MsgBox strMensajeError, vbCritical, "Error en F004_Forzar_Delimitadores_en_Excel"
    
End Sub



Public Function fun803_ObtenerCarpetaExcelActual() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta donde est� ubicado el archivo Excel actual
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener ruta completa del archivo actual
    If ThisWorkbook.Path <> "" Then
        strCarpeta = ThisWorkbook.Path
    ElseIf ActiveWorkbook.Path <> "" Then
        strCarpeta = ActiveWorkbook.Path
    Else
        strCarpeta = ""
    End If
    
    fun803_ObtenerCarpetaExcelActual = strCarpeta
    Exit Function
    
ErrorHandler:
    fun803_ObtenerCarpetaExcelActual = ""
End Function
Public Function fun804_ObtenerCarpetaEnvironmentVariable(vstrEnvironmentVariable As String) As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TEMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TEMP (compatible con Excel 97+)
    strCarpeta = Environ(UCase(vstrEnvironmentVariable))
    
    fun804_ObtenerCarpetaEnvironmentVariable = strCarpeta
    Exit Function
    
ErrorHandler:
    fun804_ObtenerCarpetaEnvironmentVariable = ""
End Function

Public Function fun804_ObtenerCarpetaTemp() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TEMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TEMP (compatible con Excel 97+)
    strCarpeta = Environ("TEMP")
    
    fun804_ObtenerCarpetaTemp = strCarpeta
    Exit Function
    
ErrorHandler:
    fun804_ObtenerCarpetaTemp = ""
End Function

Public Function fun805_ObtenerCarpetaTmp() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %TMP%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno TMP (compatible con Excel 97+)
    strCarpeta = Environ("TMP")
    
    fun805_ObtenerCarpetaTmp = strCarpeta
    Exit Function
    
ErrorHandler:
    fun805_ObtenerCarpetaTmp = ""
End Function

Public Function fun806_ObtenerCarpetaUserProfile() As String

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************

    '--------------------------------------------------------------------------
    ' Obtiene la carpeta de la variable de entorno %USERPROFILE%
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim strCarpeta As String
    
    ' Obtener variable de entorno USERPROFILE (compatible con Excel 97+)
    strCarpeta = Environ("USERPROFILE")
    
    fun806_ObtenerCarpetaUserProfile = strCarpeta
    Exit Function
    
ErrorHandler:
    fun806_ObtenerCarpetaUserProfile = ""
End Function

Public Function fun807_ValidarCarpeta(ByVal strCarpeta As String) As Boolean

    '******************************************************************************
    ' FUNCIONES AUXILIARES PARA OBTENCI�N DE CARPETAS DE RESPALDO
    ' FECHA CREACI�N: 2025-06-01
    ' AUTOR: david-joaquin-corredera-de-colsa
    ' COMPATIBILIDAD: Excel 97, 2003, 365
    '******************************************************************************
    
    '--------------------------------------------------------------------------
    ' Valida si una carpeta existe y es accesible
    '--------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim objFSO As Object
    Dim blnResultado As Boolean
    
    blnResultado = False
    
    ' Verificar que la carpeta no est� vac�a
    If Len(Trim(strCarpeta)) = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Crear objeto FileSystemObject (compatible con Excel 97+)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si la carpeta existe y es accesible
    If objFSO.FolderExists(strCarpeta) Then
        blnResultado = True
    End If
    
    Set objFSO = Nothing
    fun807_ValidarCarpeta = blnResultado
    Exit Function
    
ErrorHandler:
    Set objFSO = Nothing
    fun807_ValidarCarpeta = False
End Function


