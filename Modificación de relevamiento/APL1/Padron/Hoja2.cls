VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)

'preparacion previa al incio de funcion
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'declaracion de variables
Dim codigo As String
Dim poblacion As String


'entra si se hace doble click entre L11 y L414
If Not Intersect(target, Range("L11:L414")) Is Nothing Then

    'a codigo le otorgo el valor de del codigo de la prestacion
    'de la misma fila donde se hace doble click
    codigo = Cells(target.Row, target.Column - 3).Value
    codigoDobleClick = codigo
    
    If (codigo <> "") Then
    
    
        'asignacion de valores a variable globales (fila y columna)
        filaDobleClick = target.Row
        columnaDobleClick = target.Column
        
        'llamadas a funcion
        Call userForm_padron_copiar_datos_fijos(target)
        Call userForm_padron_copiar_datos_relevamiento(target)
        Call userForm_padron_bloquear
    
        'llamada a funcion
        Call userForm_padron_permitir_campos_requeridos(codigo)
        
        auxiliarGuardado = 0
        
        If (error = 0) Then
        
        UserForm_padron.Show
        
        End If
        
        error = 0
        
    Else
    
    MsgBox ("El campo de codigo declarado se encuentra vacio. Por favor ponerse en contacto con equipo central")
        
    End If

End If


'preparacion para el cierre de la funcion
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub


