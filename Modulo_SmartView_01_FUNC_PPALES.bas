Attribute VB_Name = "Modulo_SmartView_01_FUNC_PPALES"
Option Explicit



Public Function SmartView_CreateConnection() As Boolean
        
    'Declaracion de Variables para conectar
    Dim vUsername As String
    Dim vPassword As String
    
    ' Variables para control de errores
    Dim strFuncion As String
    ' Inicialización
    strFuncion = "SmartView_CreateConnection"
    

    ' Validar/Crear hoja UserName
    If Not fun802_SheetExists(CONST_HOJA_USERNAME) Then
        If Not F002_Crear_Hoja(CONST_HOJA_USERNAME) Then
            Err.Raise ERROR_BASE_IMPORT + 3, strFuncion, _
                "Error al crear la hoja " & gstrHoja_UserName
        End If
    End If
    ' Verificar si debemos ocultar la hoja UserName (comprobando la constante global CONST_OCULTAR_HOJA_USERNAME)
    If CONST_OCULTAR_HOJA_USERNAME = True Then
        ' Ocultar la hoja de delimitadores
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CONST_HOJA_USERNAME)
        If Not fun809_OcultarHojaDelimitadores(ws) Then
            Debug.Print "ADVERTENCIA: Error al ocultar la hoja " & gstrHoja_UserName & " - Función: F000_Comprobaciones_Iniciales - " & Now()
            ' Nota: No es un error crítico, el proceso puede continuar
        End If
    End If
    'Tomamos el vUsername de la celda Cells(2,2) de la hoja Username
    vUsername = ws.Cells(2, 2).Value
    ThisWorkbook.Save
    

    'Connection Inputs
    'Inicializamos con valor en blanco los 2 TextBox del UserForm (el de Username y el de Password)
    'UserForm_Username_Password.TextBox_Username.Value = ""
    UserForm_Username_Password.TextBox_Username.Value = vUsername
    UserForm_Username_Password.TextBox_Password.Value = ""
    'Mostramos el User Form, para pedir el Username y la Password
    UserForm_Username_Password.Show vbModal
    
    'Ahora a las variables de vUsername y vPassword le pasamos el valor de los TextBox de Username y Password
    vUsername = UserForm_Username_Password.TextBox_Username.Value
    vPassword = UserForm_Username_Password.TextBox_Password.Value
    'vUsername = InputBox("Introduzca su 'UserName' de HFM", "UserName")
    'vPassword = InputBox("Introduzca la 'Password' de HFM para el usuario " & vUsername, "Password")
    ws.Cells(2, 2).Value = vUsername
    ThisWorkbook.Save

    'Declaracion de Variables para recoger el código de error retornado por SmartView
    Dim vExisteConexion_Return As Boolean
    Dim vEliminarConexion_Return As Integer
    Dim vCrearConexion_Return As Integer
    Dim vDesconectarTodo_Return As Integer
    Dim vConexionActiva_Return As Integer
    Dim vConectar_Return As Variant 'Integer
    
    
       
    vExisteConexion_Return = HypConnectionExists(CONST_FRIENDLY_NAME)
    'Verificar si la conexion existe
      If vExisteConexion_Return Then
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox ("La conexion ya existe. Vamos a actualizarla.")
      Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox "La conexion NO existe. Vamos a crearla." & vbCrLf & "Error Number = " & vExisteConexion_Return
      End If
    
    'Desconectar de todas las aplicaciones
    If vExisteConexion_Return Then
        vDesconectarTodo_Return = HypDisconnectAll()
    Else
        vDesconectarTodo_Return = HypDisconnectAll()
    End If
    
    'Verificar si se desconecto de todas las aplicaciones
      If vDesconectarTodo_Return = 0 Then
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox ("Se desconecto correctamente de todas las aplicaciones")
      Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox "No consiguio desconectarse correctamente de todas las aplicaciones" & vbCrLf & "Error Number = " & vDesconectarTodo_Return
      End If
    
    'Eliminar la conexion CONST_FRIENDLY_NAME
      If vDesconectarTodo_Return = 0 And vExisteConexion_Return Then
        vEliminarConexion_Return = HypRemoveConnection(CONST_FRIENDLY_NAME)
      End If
    
    'Verificar si la conexion se ha eliminado
      If vEliminarConexion_Return = 0 And vExisteConexion_Return Then
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox ("Se eliminó la conexion " & CONST_FRIENDLY_NAME & " para volver a crearla con valores actualizados")
      ElseIf vExisteConexion_Return Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox "Hubo un error al intentar eliminar la conexión " & CONST_FRIENDLY_NAME & _
            " para volver a crearla con valores actualizados." & vbCrLf & "Error Number = " & vEliminarConexion_Return
      Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox "No existia la conexión " & CONST_FRIENDLY_NAME & ". " & vbCrLf & "Asi que no hubo que eliminarla." & vbCrLf & _
            "Vamos a crearla con valores actualizados." & vbCrLf & "Error Number = " & vEliminarConexion_Return
      End If
    
    'Crear una conexion via SmartView a la aplicacion deseada
      If vEliminarConexion_Return = 0 Then
        vCrearConexion_Return = HypCreateConnection(Empty, vUsername, vPassword, CONST_PROVIDER, CONST_PROVIDER_URL, CONST_SERVER_NAME, _
               CONST_APPLICATION_NAME, CONST_DATABASE_NAME, CONST_FRIENDLY_NAME, CONST_DESCRIPTION)
      Else
        vCrearConexion_Return = HypCreateConnection(Empty, vUsername, vPassword, CONST_PROVIDER, CONST_PROVIDER_URL, CONST_SERVER_NAME, _
               CONST_APPLICATION_NAME, CONST_DATABASE_NAME, CONST_FRIENDLY_NAME, CONST_DESCRIPTION)
      End If
      
    'Verificar si la conexion se ha creado
      If vCrearConexion_Return = 0 Then
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox ("Se creo la conexion" & CONST_FRIENDLY_NAME & " correctamente, con valores actualizados")
      Else
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox "Fallo la creacion de la conexion" & CONST_FRIENDLY_NAME & "." & vbCrLf & "Error Number = " & vCrearConexion_Return
      End If

    'Conectar a SmartView usando la nueva conexion
    If vCrearConexion_Return = 0 Then
        vConectar_Return = HypConnect(Empty, vUsername, vPassword, CONST_FRIENDLY_NAME)
    End If
    
    'Verificar si nos hemos conectado
      If vConectar_Return = 0 Then
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox ("Nos hemos conectado correctamente a " & CONST_FRIENDLY_NAME)
      Else
         If CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION Then MsgBox "Fallo al conectarnos." & vbCrLf & "Error Number = " & vConectar_Return
      End If
    
    'Fijar como conexion activa la nueva conexion
    If vConectar_Return = 0 Then
        vConexionActiva_Return = HypSetActiveConnection(CONST_FRIENDLY_NAME)
    End If

    'Verificar si hemos puesto como conexion activa la recien creada
      If vConexionActiva_Return = 0 Then
         MsgBox ("Hemos establecido " & CONST_FRIENDLY_NAME & " como conexion activa")
      Else
         MsgBox "Fallo al establecer " & CONST_FRIENDLY_NAME & " como conexion activa." & vbCrLf & "Error Number = " & vConexionActiva_Return
      End If



    If vConexionActiva_Return = 0 Then
        SmartView_CreateConnection = True
    Else
        SmartView_CreateConnection = False
    End If
    

End Function

Public Function SmartView_Options_DataOptions_Estandar(vHoja As Variant) As Boolean

    Dim vInteger01, vInteger02, vInteger03, vInteger04, vInteger05, vInteger06, vInteger07, vInteger08 As Double
    
    vInteger01 = SmartView_Options_MemberOptions_Indent_None(vHoja)
    
    vInteger02 = SmartView_Options_MemberOptions_DisplayNameOnly(vHoja)
    
    vInteger03 = SmartView_Options_DataOptions_CellDisplay(vHoja)
    
    vInteger04 = SmartView_Options_DataOptions_Supress_Missing(vHoja, False)
    
    vInteger05 = SmartView_Options_DataOptions_Supress_Zero(vHoja, False)
    
    vInteger06 = SmartView_Options_DataOptions_Supress_Repeated(vHoja, False)
    
    vInteger07 = SmartView_Options_DataOptions_Supress_Invalid(vHoja, False)
    
    vInteger08 = SmartView_Options_DataOptions_Supress_NoAccess(vHoja, False)
    
    If (vInteger01 <> 0) Or (vInteger02 <> 0) Or (vInteger03 <> 0) Or (vInteger04 <> 0) Or (vInteger05 <> 0) Or (vInteger06 <> 0) Or _
        (vInteger07 <> 0) Or (vInteger08 <> 0) Then
        SmartView_Options_DataOptions_Estandar = False
    Else
        SmartView_Options_DataOptions_Estandar = True
    End If

    
End Function

Public Function SmartView_Retrieve(vHoja As Variant) As Boolean
    Dim vLong As Long
    
    ThisWorkbook.Worksheets(vHoja).Activate
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    vHoja = ThisWorkbook.ActiveSheet.Name
    vLong = HypRetrieve(vHoja)
    'MsgBox "vLong = " & vLong
    
    If vLong = 0 Then
        MsgBox "Refrescado con exito."
    Else
        MsgBox "Hubo un error al refrescar. Error Number = " & SmartView_Retrieve
    End If
    If vLong = 0 Then
        SmartView_Retrieve = True
    Else
        SmartView_Retrieve = False
    End If
    
End Function

Public Function SmartView_Submit(vHoja As Variant) As Boolean
    Dim vLong As Long
    
    ThisWorkbook.Worksheets(vHoja).Activate
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    vHoja = ThisWorkbook.ActiveSheet.Name
    vLong = HypSubmitData(vHoja)
    'MsgBox "vLong = " & vLong
    
    If vLong = 0 Then
        MsgBox "Submit ejecutado con exito."
    Else
        MsgBox "Hubo un error al ejecutar Submit. Error Number = " & SmartView_Submit
    End If
    If vLong = 0 Then
        SmartView_Submit = True
    Else
        SmartView_Submit = False
    End If
    
End Function

Public Function SmartView_Submit_without_Refresh(vHoja As Variant) As Boolean
    Dim vLong As Long
    
    ThisWorkbook.Worksheets(vHoja).Activate
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    vHoja = ThisWorkbook.ActiveSheet.Name
    vLong = HypSubmitSelectedRangeWithoutRefresh(vHoja, False, False, False)
    'MsgBox "vLong = " & vLong
    
    If vLong = 0 Then
        MsgBox "Submitted without Refresh - ejecutado con exito."
    Else
        MsgBox "Hubo un error al ejecutar - Submit without Refresh. Error Number = " & SmartView_Submit_without_Refresh
    End If
    If vLong = 0 Then
        SmartView_Submit_without_Refresh = True
    Else
        SmartView_Submit_without_Refresh = False
    End If
    
End Function

