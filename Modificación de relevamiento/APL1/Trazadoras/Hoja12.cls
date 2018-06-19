VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja12"
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
    filaDobleClicktz12 = Target.Row
    columnaDobleClickTz12 = Target.Column

    'llamadas a funcion
    Call copiar_tz12_datos_fijos(filaDobleClicktz12)
    Call userForm_tz12_copiar_datos_relevamiento(filaDobleClicktz12)
    Call userForm_tz12_bloquear

    'llamada a funcion
    Call userForm_tz12_permitir_campos_requeridos
    
    
    userForm_tz12.Show
    
    
End If


'preparacion para el cierre de la funcion
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub



