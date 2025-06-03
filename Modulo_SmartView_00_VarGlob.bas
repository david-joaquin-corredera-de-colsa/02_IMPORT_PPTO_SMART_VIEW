Attribute VB_Name = "Modulo_SmartView_00_VarGlob"
Option Explicit

'******************************************************************************
' Módulo: SV_Global_Variables_y_Constantes
' Fecha y Hora de Creación: 2025-05-26 10:04:46 UTC
' Autor: david-joaquin-corredera-de-colsa
'
' Descripción:
' Este módulo contiene todas las variables y constantes globales utilizadas en el sistema para SmartView
'******************************************************************************

' Constantes para mostrar o no mensajes durante la ejecución
Public Const CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS As Boolean = False
Public Const CONST_MOSTRAR_MENSAJES_SMARTVIEW_CREAR_CONEXION As Boolean = False

' Constantes para la Conexion
Public Const CONST_PROVIDER As String = "Hyperion Financial Management"
Public Const CONST_PROVIDER_URL As String = "http://sv3572.logista.local:19000/hfmadf/officeprovider"
Public Const CONST_SERVER_NAME As String = "HFM"
Public Const CONST_APPLICATION_NAME As String = "BUCONS1012"
Public Const CONST_DATABASE_NAME As String = "BUCONS1012"
Public Const CONST_FRIENDLY_NAME As String = "Conexion_BUCONS1012_Presupuesto"
Public Const CONST_DESCRIPTION As String = "Conexion_BUCONS1012_Presupuesto"

' Constantes para los SmartView > Options > Data Options
Public Const CONST_INDENT_SETTING As Integer = 5
    Public Const CONST_INDENT_NONE As Integer = 0
    Public Const CONST_INDENT_CHILD As Integer = 1
    Public Const CONST_INDENT_PARENT As Integer = 2
Public Const CONST_SUPRESS_MISSING_SETTING As Integer = 6
Public Const CONST_SUPRESS_ZERO_SETTING As Integer = 7
Public Const CONST_ENABLE_NOACCESS_MEMBERS_SETTING As Integer = 9
Public Const CONST_ENABLE_REPEATED_MEMBERS_SETTING As Integer = 10
Public Const CONST_ENABLE_INVALID_MEMBERS_SETTING As Integer = 11
Public Const CONST_CELL_DISPLAY_SETTING As Integer = 15
    Public Const CONST_CELL_DISPLAY_SHOW_DATA As Integer = 0
    Public Const CONST_CELL_DISPLAY_SHOW_CALC_STATUS As Integer = 1
    Public Const CONST_CELL_DISPLAY_SHOW_PROCESS_MANAGEMENT As Integer = 2
Public Const CONST_DISPLAY_MEMBER_NAME_SETTING As Integer = 16
    Public Const CONST_DISPLAY_NAME_ONLY As Integer = 0
    Public Const CONST_DISPLAY_AND_DESCRIPTION As Integer = 1
    Public Const CONST_DISPLAY_DESCRIPTION_ONLY As Integer = 2


