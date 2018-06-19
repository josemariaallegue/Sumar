VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

'preparacion previa al incio de funcion
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


'entra si se hace doble click entre L11 y L414
If Not Intersect(Target, Range("I10:I500")) Is Nothing Then
    
    'asignacion de valores a variable globales (fila y columna)
    filaDobleClickTz13 = Target.Row
    columnaDobleClickTz13 = Target.Column
    
    'llamadas a funcion
    Call copiar_tz13_datos_fijos(filaDobleClickTz13)
    Call userForm_tz13_copiar_datos_relevamiento(filaDobleClickTz13)
    Call userForm_tz13_bloquear

    'llamada a funcion
    Call userForm_tz13_permitir_campos_requeridos
    
    'If (Error = 0) Then
    
    userForm_tz13.Show
    
    'End If
    
    'Error = 0
    

End If


'preparacion para el cierre de la funcion
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub



