Attribute VB_Name = "Modulo_SmartView_03_FUNC_AUX"
Option Explicit


Public Function SmartView_Options_MemberOptions_Indent_None(vNombreHoja As Variant) As Long
    SmartView_Options_MemberOptions_Indent_None = HypSetSheetOption(vNombreHoja, CONST_INDENT_SETTING, CONST_INDENT_NONE)
    If SmartView_Options_MemberOptions_Indent_None = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Member Options > Indent = None"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Member Options > Indent = None." & vbCrLf & "Error Number = " & SmartView_Options_MemberOptions_Indent_None
    End If
End Function

Public Function SmartView_Options_DataOptions_Supress_Missing(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_Supress_Missing = HypSetSheetOption(vNombreHoja, CONST_SUPRESS_MISSING_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_Supress_Missing = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Missing = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Data Options > Supress Missing = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_Supress_Missing
    End If
    
End Function
Public Function SmartView_Options_DataOptions_Supress_Zero(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_Supress_Zero = HypSetSheetOption(vNombreHoja, CONST_SUPRESS_ZERO_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_Supress_Zero = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Zero = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Zero = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_Supress_Zero
    End If
    
End Function

Public Function SmartView_Options_DataOptions_Supress_Repeated(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_Supress_Repeated = HypSetSheetOption(vNombreHoja, CONST_ENABLE_REPEATED_MEMBERS_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_Supress_Repeated = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Repeated = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Repeated = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_Supress_Repeated
    End If
End Function
Public Function SmartView_Options_DataOptions_Supress_Invalid(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_Supress_Invalid = HypSetSheetOption(vNombreHoja, CONST_ENABLE_INVALID_MEMBERS_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_Supress_Invalid = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Invalid = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress Invalid = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_Supress_Invalid
    End If
    
End Function
Public Function SmartView_Options_DataOptions_Supress_NoAccess(vNombreHoja As Variant, vTrueFalse As Boolean) As Long
    SmartView_Options_DataOptions_Supress_NoAccess = HypSetSheetOption(vNombreHoja, CONST_ENABLE_NOACCESS_MEMBERS_SETTING, vTrueFalse)
    If SmartView_Options_DataOptions_Supress_NoAccess = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress NoAccess = False"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > Supress NoAccess = False." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_Supress_NoAccess
    End If
End Function

Public Function SmartView_Options_DataOptions_CellDisplay(vNombreHoja As Variant) As Long
    SmartView_Options_DataOptions_CellDisplay = HypSetSheetOption(vNombreHoja, CONST_CELL_DISPLAY_SETTING, CONST_CELL_DISPLAY_SHOW_DATA)
    If SmartView_Options_DataOptions_CellDisplay = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Data Options > Cell Display = Data"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Data Options > Cell Display = Data." & vbCrLf & "Error Number = " & SmartView_Options_DataOptions_CellDisplay
    End If
End Function

Public Function SmartView_Options_MemberOptions_DisplayNameOnly(vNombreHoja As Variant) As Long
    SmartView_Options_MemberOptions_DisplayNameOnly = HypSetSheetOption(vNombreHoja, CONST_DISPLAY_MEMBER_NAME_SETTING, CONST_DISPLAY_NAME_ONLY)
    If SmartView_Options_MemberOptions_DisplayNameOnly = 0 Then
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Se establecio con exito la opcion " & vbCrLf & "SmartView > Options > Member Options > Member Display = Name Only"
    Else
        If CONST_MOSTRAR_MENSAJES_SMARTVIEW_OPTIONS Then MsgBox "Hubo un error al intentar establecer la opcion " & vbCrLf & "SmartView > Options > Member Options > Member Display = Name Only." & vbCrLf & "Error Number = " & SmartView_Options_MemberOptions_DisplayNameOnly
    End If
End Function


