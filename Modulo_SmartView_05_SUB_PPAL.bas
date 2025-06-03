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

Sub M002_SmartView_Paso_01_v2(vNombreDeLaHoja As String)

    Dim x As Boolean
    x = SmartView_CreateConnection
    
    
    'Dim vNombreDeLaHoja As String
    'vNombreDeLaHoja = ThisWorkbook.ActiveSheet.Name
    
    Dim vReturn_SmartView_Options_DataOptions As Integer
    vReturn_SmartView_Options_DataOptions = SmartView_Options_DataOptions_Estandar(vNombreDeLaHoja)
    
    Dim vReturn_SmartView_Retrieve As Integer
    vReturn_SmartView_Retrieve = SmartView_Retrieve(vNombreDeLaHoja)
    
End Sub

