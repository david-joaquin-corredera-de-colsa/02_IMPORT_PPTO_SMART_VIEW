Attribute VB_Name = "Modulo_521_FUNC_Auxiliares_03"

Option Explicit

Public Function fun812_CopiarContenidoCompleto(ByRef wsOrigen As Worksheet, _
                                               ByRef wsDestino As Worksheet) As Boolean
    
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR CORREGIDA: fun812_CopiarContenidoCompleto
    ' Fecha y Hora de Modificaci�n: 2025-06-01 19:34:00 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '
    ' Descripci�n:
    ' Copia todo el contenido de una hoja de trabajo a otra hoja de destino
    ' MANTENIENDO LA POSICI�N ORIGINAL de los datos (ej: si origen est� en B2,
    ' destino tambi�n estar� en B2).
    '******************************************************************************
    On Error GoTo GestorErrores
    
    Dim rngUsedOrigen As Range
    Dim strCeldaDestino As String
    
    ' Limpiar hoja destino
    If Not fun801_LimpiarHoja(wsDestino.Name) Then
        fun812_CopiarContenidoCompleto = False
        Exit Function
    End If
    
    ' Verificar que hay contenido en la hoja origen
    If wsOrigen.UsedRange Is Nothing Then
        fun812_CopiarContenidoCompleto = True
        Exit Function
    End If
    
    ' Obtener rango usado de origen
    Set rngUsedOrigen = wsOrigen.UsedRange
    
    ' Calcular celda destino manteniendo posici�n original
    ' Si el rango origen empieza en B2, el destino tambi�n empezar� en B2
    strCeldaDestino = wsDestino.Cells(rngUsedOrigen.Row, rngUsedOrigen.Column).Address
    
    ' Copiar manteniendo posici�n original
    rngUsedOrigen.Copy wsDestino.Range(strCeldaDestino)
    Application.CutCopyMode = False
    
    fun812_CopiarContenidoCompleto = True
    Exit Function
    
GestorErrores:
    Application.CutCopyMode = False
    fun812_CopiarContenidoCompleto = False
End Function


Public Function fun813_DetectarRangoCompleto(ByRef ws As Worksheet, _
                                            ByRef vFila_Inicial As Long, _
                                            ByRef vFila_Final As Long, _
                                            ByRef vColumna_Inicial As Long, _
                                            ByRef vColumna_Final As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun813_DetectarRangoCompleto
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim rngUsado As Range
    
    ' Obtener rango usado
    Set rngUsado = ws.UsedRange
    
    If rngUsado Is Nothing Then
        vFila_Inicial = 0
        vFila_Final = 0
        vColumna_Inicial = 0
        vColumna_Final = 0
        fun813_DetectarRangoCompleto = False
        Exit Function
    End If
    
    ' Detectar rangos
    vFila_Inicial = rngUsado.Row
    vFila_Final = rngUsado.Row + rngUsado.Rows.Count - 1
    vColumna_Inicial = rngUsado.Column
    vColumna_Final = rngUsado.Column + rngUsado.Columns.Count - 1
    
    fun813_DetectarRangoCompleto = True
    Exit Function
    
GestorErrores:
    vFila_Inicial = 0
    vFila_Final = 0
    vColumna_Inicial = 0
    vColumna_Final = 0
    fun813_DetectarRangoCompleto = False
End Function


Public Sub fun814_MostrarInformacionColumnas(ByVal vColumna_Inicial As Long, _
                                            ByVal vColumna_Final As Long, _
                                            ByVal vColumna_IdentificadorDeLinea As Long, _
                                            ByVal vColumna_LineaRepetida As Long, _
                                            ByVal vColumna_LineaTratada As Long, _
                                            ByVal vColumna_LineaSuma As Long, _
                                            ByVal vFila_Inicial As Long, _
                                            ByVal vFila_Final As Long)
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun814_MostrarInformacionColumnas
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    Dim strMensaje As String
    
    strMensaje = "INFORMACI�N DE VARIABLES DE COLUMNAS DE CONTROL" & vbCrLf & vbCrLf & _
                 "RANGOS DETECTADOS:" & vbCrLf & _
                 "- Fila Inicial: " & vFila_Inicial & vbCrLf & _
                 "- Fila Final: " & vFila_Final & vbCrLf & _
                 "- Columna Inicial: " & vColumna_Inicial & vbCrLf & _
                 "- Columna Final: " & vColumna_Final & vbCrLf & vbCrLf & _
                 "COLUMNAS DE CONTROL CALCULADAS:" & vbCrLf & _
                 "- vColumna_IdentificadorDeLinea = " & vColumna_IdentificadorDeLinea & _
                 " (Inicial+" & (vColumna_IdentificadorDeLinea - vColumna_Inicial) & ")" & vbCrLf & _
                 "- vColumna_LineaRepetida = " & vColumna_LineaRepetida & _
                 " (Inicial+" & (vColumna_LineaRepetida - vColumna_Inicial) & ")" & vbCrLf & _
                 "- vColumna_LineaTratada = " & vColumna_LineaTratada & _
                 " (Inicial+" & (vColumna_LineaTratada - vColumna_Inicial) & ")" & vbCrLf & _
                 "- vColumna_LineaSuma = " & vColumna_LineaSuma & _
                 " (Inicial+" & (vColumna_LineaSuma - vColumna_Inicial) & ")" & vbCrLf & vbCrLf & _
                 "Para desactivar este mensaje, cambiar True por False en el c�digo."
    
    MsgBox strMensaje, vbInformation, "Variables de Columnas de Control"
End Sub


Public Function fun815_BorrarColumnasInnecesarias(ByRef ws As Worksheet, _
                                                  ByVal vFila_Inicial As Long, _
                                                  ByVal vFila_Final As Long, _
                                                  ByVal vColumna_Inicial As Long, _
                                                  ByVal vColumna_IdentificadorDeLinea As Long, _
                                                  ByVal vColumna_LineaRepetida As Long, _
                                                  ByVal vColumna_LineaSuma As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun815_BorrarColumnasInnecesarias
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim i As Long
    
    ' Borrar columna identificador de l�nea
    ws.Range(ws.Cells(vFila_Inicial, vColumna_IdentificadorDeLinea), _
             ws.Cells(vFila_Final, vColumna_IdentificadorDeLinea)).Clear
    
    ' Borrar columna l�nea repetida
    ws.Range(ws.Cells(vFila_Inicial, vColumna_LineaRepetida), _
             ws.Cells(vFila_Final, vColumna_LineaRepetida)).Clear
    
    ' Borrar columnas a la izquierda de vColumna_Inicial (excluyendo vColumna_Inicial)
    If vColumna_Inicial > 1 Then
        For i = 1 To vColumna_Inicial - 1
            ws.Range(ws.Cells(vFila_Inicial, i), _
                     ws.Cells(vFila_Final, i)).Clear
        Next i
    End If
    
    ' Borrar columnas a la derecha de vColumna_LineaSuma (excluyendo vColumna_LineaSuma)
    For i = vColumna_LineaSuma + 1 To ws.Columns.Count
        ' Solo limpiar si hay contenido para optimizar rendimiento
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(vFila_Inicial, i), _
                                                         ws.Cells(vFila_Final, i))) > 0 Then
            ws.Range(ws.Cells(vFila_Inicial, i), _
                     ws.Cells(vFila_Final, i)).Clear
        Else
            Exit For ' Si no hay contenido, salir del bucle
        End If
    Next i
    
    fun815_BorrarColumnasInnecesarias = True
    Exit Function
    
GestorErrores:
    fun815_BorrarColumnasInnecesarias = False
End Function


Public Function fun816_FiltrarLineasEspecificas(ByRef ws As Worksheet, _
                                               ByVal vFila_Inicial As Long, _
                                               ByVal vFila_Final As Long, _
                                               ByVal vColumna_Inicial As Long, _
                                               ByVal vColumna_LineaTratada As Long) As Boolean
    
    '******************************************************************************
    ' FUNCI�N AUXILIAR: fun816_FiltrarLineasEspecificas
    ' Fecha y Hora de Creaci�n: 2025-06-01 19:20:05 UTC
    ' Autor: david-joaquin-corredera-de-colsa
    '******************************************************************************
    
    On Error GoTo GestorErrores
    
    Dim i As Long
    Dim vValor_Columna_Inicial As String
    Dim vValor_Primer_Caracter_Columna_Inicial As String
    Dim vValor_Columna_LineaTratada As String
    Dim blnBorrarLinea As Boolean
    
    ' Recorrer l�neas desde la final hacia la inicial para evitar problemas de �ndices
    For i = vFila_Final To vFila_Inicial Step -1
        
        ' Reinicializar variables para cada l�nea
        vValor_Columna_Inicial = ""
        vValor_Primer_Caracter_Columna_Inicial = ""
        vValor_Columna_LineaTratada = ""
        blnBorrarLinea = False
        
        ' Obtener valor de la primera columna
        vValor_Columna_Inicial = Trim(CStr(ws.Cells(i, vColumna_Inicial).Value))
        
        ' Obtener primer car�cter si hay contenido
        If Len(vValor_Columna_Inicial) > 0 Then
            vValor_Primer_Caracter_Columna_Inicial = Left(vValor_Columna_Inicial, 1)
        Else
            vValor_Primer_Caracter_Columna_Inicial = ""
        End If
        
        ' Obtener valor de columna l�nea tratada
        vValor_Columna_LineaTratada = Trim(CStr(ws.Cells(i, vColumna_LineaTratada).Value))
        
        ' Evaluar criterios para borrar l�nea
        If (vValor_Primer_Caracter_Columna_Inicial = "!") Or _
           (vValor_Columna_Inicial = "") Or _
           (Len(Trim(vValor_Columna_Inicial)) = 0) Or _
           (vValor_Columna_LineaTratada = CONST_TAG_LINEA_TRATADA) Then
            
            blnBorrarLinea = True
        End If
        
        ' Borrar contenido de toda la l�nea si cumple criterios
        If blnBorrarLinea Then
            ws.Rows(i).ClearContents
        End If
        
    Next i
    
    fun816_FiltrarLineasEspecificas = True
    Exit Function
    
GestorErrores:
    fun816_FiltrarLineasEspecificas = False
End Function


