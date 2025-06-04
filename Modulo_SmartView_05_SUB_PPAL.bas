Attribute VB_Name = "Modulo_SmartView_05_SUB_PPAL"
Option Explicit

Sub M002_SmartView_Paso_01()

    Dim x As Boolean
    x = SmartView_CreateConnection
    
    
    Dim vNombreDeLaHoja As String
    vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
    Dim vReturn_SmartView_Options_DataOptions As Integer
    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Retrieve As Integer
    vReturn_SmartView_Retrieve = SmartView_Retrieve(vNombreDeLaHoja)
    
End Sub

Sub M002_SmartView_Paso_01_CrearConexiones_EstablecerOpciones_CrearAdHoc(vNombreDeLaHoja As String)

    Dim x As Boolean
    x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
    Dim vReturn_SmartView_Options_DataOptions As Integer
    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Retrieve As Integer
    vReturn_SmartView_Retrieve = SmartView_Retrieve(vNombreDeLaHoja)
    
End Sub

Sub M003_SmartView_Paso_02_Submit(vNombreDeLaHoja As String)

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit As Integer
    vReturn_SmartView_Submit = SmartView_Submit(vNombreDeLaHoja)
    
End Sub

Sub M003_SmartView_Paso_02_Submit_without_Refresh(vNombreDeLaHoja As String)

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit_without_Refresh As Integer
    vReturn_SmartView_Submit_without_Refresh = SmartView_Submit_without_Refresh(vNombreDeLaHoja)
    
End Sub

Public Sub xx_Stand_Alone_M003_SmartView_Paso_02_Submit_without_Refresh()

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    Dim vNombreDeLaHoja As String
    vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    MsgBox "vNombreDeLaHoja=" & vNombreDeLaHoja
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit_without_Refresh As Integer
    vReturn_SmartView_Submit_without_Refresh = SmartView_Submit_without_Refresh(vNombreDeLaHoja)
    
End Sub
Sub xx_Stand_Alone_M003_SmartView_Paso_02_Submit()

'    Dim x As Boolean
'    x = SmartView_CreateConnection
    
    
    Dim vNombreDeLaHoja As String
    vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    MsgBox "vNombreDeLaHoja=" & vNombreDeLaHoja
    
'    Dim vReturn_SmartView_Options_DataOptions As Integer
'    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Submit As Integer
    vReturn_SmartView_Submit = SmartView_Submit(vNombreDeLaHoja)
    
End Sub

Public Sub xx_Editar_Celdas()
    Dim r As Integer
    Dim c As Integer
    Dim vValor As Variant
    
    
    For r = 8 To 10
        For c = 13 To 24
            vValor = Cells(r, c).Value
            Cells(r, c).Value = vValor
        Next c
    Next r
End Sub
