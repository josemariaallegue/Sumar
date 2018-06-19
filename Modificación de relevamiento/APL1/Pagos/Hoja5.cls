VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

'preparacion previa al incio de funcion
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'declaracion de variables
Dim codigo As String
Dim poblacion As String


'entra si se hace doble click entre L11 y L414
If Not Intersect(Target, Range("M11:M399")) Is Nothing Then
    
    'asignacion de valores a variable globales (fila y columna)
    filaDobleClickNI = Target.Row
    columnaDobleClickNI = Target.Column

    'a codigo le otorgo el valor de del codigo de la prestacion
    'de la misma fila donde se hace doble click
    codigo = Cells(Target.Row, Target.Column - 3).Value
    codigoDobleClickNI = codigo

    'llamadas a funcion
    Call copiar_ni_datos_fijos(Target)
    Call userForm_ni_copiar_datos_relevamiento(Target)
    Call userForm_ni_bloquear

    'llamada a funcion
    Call userForm_ni_permitir_campos_requeridos(codigo)

    If (errorNI = 0) Then
    
    userForm_ni.Show
    
    End If

    errorNI = 0
    

End If


'preparacion para el cierre de la funcion
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
