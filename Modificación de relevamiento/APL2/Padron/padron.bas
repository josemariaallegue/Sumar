Attribute VB_Name = "padron"
'declaracion de variables globales

Public filaDobleClick As Integer

Public columnaDobleClick As Integer

Public codigoDobleClick As String

Public auxiliarTarget As Range

Public estudiosFlag  As Integer

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistente As Integer

'existe para comprobar si se guardo o se hizo un cambio al salir
'0 es el valor inicial - no hubo cambios
'1 se entro a algun change de los textboxs o comoboboxs pero no se guardo
'2 significa que se guardo
Public auxiliarGuardado As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public error As Integer


'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub userForm_padron_copiar_datos_fijos(ByVal target As Range)

UserForm_padron.TextBox_n_efector.Text = Cells(target.Row, target.Column - 10).Value
UserForm_padron.TextBox_denominacion_efector.Text = Cells(target.Row, target.Column - 9).Value
UserForm_padron.TextBox_beneficiario.Text = Cells(target.Row, target.Column - 8).Value & " " & Cells(target.Row, target.Column - 7).Value
UserForm_padron.TextBox_documento.Text = Cells(target.Row, target.Column - 6).Value
UserForm_padron.TextBox_clave_beneficiario.Text = Cells(target.Row, target.Column - 5).Value
UserForm_padron.TextBox_codigo.Text = Cells(target.Row, target.Column - 3).Value
UserForm_padron.TextBox_fecha_prestacion.Text = Cells(target.Row, target.Column - 1).Value
UserForm_padron.TextBox_fecha_nacimiento.Text = Cells(target.Row, target.Column + 4).Value
UserForm_padron.TextBox_edad.Text = Cells(target.Row, target.Column + 20).Value
UserForm_padron.TextBox_linea_cuidado.Text = Cells(target.Row, target.Column - 2).Value

'verifica que si la celda de nombre de prestacion contiene "La prestacion no corresponde al grupo poblacional" si
'esto es verdadero le otorga un valor al textbox de control de fuente de informacion
If (Cells(filaDobleClick, columnaDobleClick - 2).Value = "La prestación no corresponde al grupo poblacional") Then
UserForm_padron.dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"
End If

End Sub


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_padron_bloquear()

UserForm_padron.TextBox_n_efector.Locked = True
UserForm_padron.TextBox_denominacion_efector.Locked = True
UserForm_padron.TextBox_beneficiario.Locked = True
UserForm_padron.TextBox_clave_beneficiario.Locked = True
UserForm_padron.TextBox_documento.Locked = True
UserForm_padron.TextBox_codigo.Locked = True
UserForm_padron.TextBox_fecha_prestacion.Locked = True
UserForm_padron.TextBox_fecha_nacimiento.Locked = True
UserForm_padron.TextBox_edad.Locked = True
UserForm_padron.TextBox_linea_cuidado.Locked = True

'userForm_padron.dato_fuente.Locked = True
'userForm_padron.dato_diagnostico.Locked = True
UserForm_padron.dato_estudios.Locked = True
UserForm_padron.dato_evaluacion_riesgo.Locked = True
UserForm_padron.dato_ta.Locked = True
UserForm_padron.dato_imc.Locked = True
UserForm_padron.dato_percentilo.Locked = True
UserForm_padron.dato_peso.Locked = True
UserForm_padron.dato_talla.Locked = True
UserForm_padron.dato_plan_seguimiento.Locked = True
UserForm_padron.dato_tratamiento_instaurado.Locked = True
UserForm_padron.dato_transcripcion.Locked = True
UserForm_padron.dato_constancia_inmunizaciones.Locked = True
UserForm_padron.dato_firma.Locked = True
UserForm_padron.dato_sello.Locked = True
UserForm_padron.dato_control_prenatal.Locked = True
'userForm_padron.dato_observaciones.Locked = True
'userForm_padron.dato_control_fuente.Locked = True
'userForm_padron.dato_validacion.Locked = True


End Function


'brief: copia los datos ya relevados a un userForm_padron
'param: rango donde se hizo doble click
'return: void

Sub userForm_padron_copiar_datos_relevamiento(ByVal target As Range)

'declaro y otorgo valor a una variable para evitar modificaciones de mas si se cambio la leyenda
Dim leyenda As String

leyenda = "Dato no obligatorio"


UserForm_padron.dato_fuente = Cells(target.Row, target.Column + 1).Value

'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
If (UserForm_padron.dato_fuente.Text <> "No consta fuente de información" And UserForm_padron.dato_fuente.Text <> "Prestación inexistente" And _
UserForm_padron.dato_fuente.Text <> "Caso duplicado" And UserForm_padron.dato_control_fuente.Text <> "Fuente invalida") Then

    UserForm_padron.dato_diagnostico.Text = Cells(target.Row, target.Column + 5).Value
    
    'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
    'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
    'lineas de codigo de mas
    
    UserForm_padron.dato_estudios.Text = Cells(target.Row, target.Column + 6).Value
    If (UserForm_padron.dato_estudios.Text = "") Then
    UserForm_padron.dato_estudios.Text = leyenda
    End If
    
    UserForm_padron.dato_evaluacion_riesgo.Text = Cells(target.Row, target.Column + 7).Value
    If (UserForm_padron.dato_evaluacion_riesgo.Text = "") Then
    UserForm_padron.dato_evaluacion_riesgo.Text = leyenda
    End If
    
    UserForm_padron.dato_ta.Text = Cells(target.Row, target.Column + 8).Value
    If (UserForm_padron.dato_ta.Text = "") Then
    UserForm_padron.dato_ta.Text = leyenda
    End If
    
    UserForm_padron.dato_imc.Text = Cells(target.Row, target.Column + 9).Value
    If (UserForm_padron.dato_imc.Text = "") Then
    UserForm_padron.dato_imc.Text = leyenda
    End If
    
    UserForm_padron.dato_percentilo.Text = Cells(target.Row, target.Column + 10).Value
    If (UserForm_padron.dato_percentilo.Text = "") Then
    UserForm_padron.dato_percentilo.Text = leyenda
    End If
    
    UserForm_padron.dato_peso.Text = Cells(target.Row, target.Column + 11).Value
    If (UserForm_padron.dato_peso.Text = "") Then
    UserForm_padron.dato_peso.Text = leyenda
    End If
    
    UserForm_padron.dato_talla.Text = Cells(target.Row, target.Column + 12).Value
    If (UserForm_padron.dato_talla.Text = "") Then
    UserForm_padron.dato_talla.Text = leyenda
    End If
    
    UserForm_padron.dato_plan_seguimiento.Text = Cells(target.Row, target.Column + 13).Value
    If (UserForm_padron.dato_plan_seguimiento.Text = "") Then
    UserForm_padron.dato_plan_seguimiento.Text = leyenda
    End If
    
    UserForm_padron.dato_tratamiento_instaurado.Text = Cells(target.Row, target.Column + 14).Value
    If (UserForm_padron.dato_tratamiento_instaurado.Text = "") Then
    UserForm_padron.dato_tratamiento_instaurado.Text = leyenda
    End If
    
    UserForm_padron.dato_transcripcion.Text = Cells(target.Row, target.Column + 15).Value
    If (UserForm_padron.dato_transcripcion.Text = "") Then
    UserForm_padron.dato_transcripcion.Text = leyenda
    End If
    
    UserForm_padron.dato_constancia_inmunizaciones.Text = Cells(target.Row, target.Column + 16).Value
    If (UserForm_padron.dato_constancia_inmunizaciones.Text = "") Then
    UserForm_padron.dato_constancia_inmunizaciones.Text = leyenda
    End If
    
    UserForm_padron.dato_firma.Text = Cells(target.Row, target.Column + 17).Value
    If (UserForm_padron.dato_firma.Text = "") Then
    UserForm_padron.dato_firma.Text = leyenda
    End If
    
    UserForm_padron.dato_sello.Text = Cells(target.Row, target.Column + 18).Value
    If (UserForm_padron.dato_sello.Text = "") Then
    UserForm_padron.dato_sello.Text = leyenda
    End If
    
    UserForm_padron.dato_control_prenatal.Text = Cells(target.Row, target.Column + 3).Value
    If (UserForm_padron.dato_control_prenatal.Text = "") Then
    UserForm_padron.dato_control_prenatal.Text = leyenda
    End If
    
    'userForm_padron.dato_control_fuente.Text = Cells(Target.Row, Target.Column + 2).Value
    
    'userForm_padron.dato_validacion.Text = Cells(Target.Row, Target.Column + 21).Value
Else

    Call userForm_padron_dato_no_obligatorio
    
End If

UserForm_padron.dato_observaciones.Text = Cells(target.Row, target.Column + 19).Value

End Sub


'brief: copia los datos del userForm_padron al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_padron_guardar_datos(ByVal fila As Integer)

Dim leyenda As String
Dim i As Integer

leyenda = "Dato no obligatorio"

Cells(fila, 13).Value = UserForm_padron.dato_fuente.Text
Cells(fila, 17).Value = UserForm_padron.dato_diagnostico.Text
Cells(fila, 18).Value = UserForm_padron.dato_estudios.Text
Cells(fila, 19).Value = UserForm_padron.dato_evaluacion_riesgo.Text
Cells(fila, 20).Value = UserForm_padron.dato_ta.Text
Cells(fila, 21).Value = UserForm_padron.dato_imc.Text
Cells(fila, 22).Value = UserForm_padron.dato_percentilo.Text
Cells(fila, 23).Value = UserForm_padron.dato_peso.Text
Cells(fila, 24).Value = UserForm_padron.dato_talla.Text
Cells(fila, 25).Value = UserForm_padron.dato_plan_seguimiento.Text
Cells(fila, 26).Value = UserForm_padron.dato_tratamiento_instaurado.Text
Cells(fila, 27).Value = UserForm_padron.dato_transcripcion.Text
Cells(fila, 28).Value = UserForm_padron.dato_constancia_inmunizaciones.Text
Cells(fila, 29).Value = UserForm_padron.dato_firma.Text
Cells(fila, 30).Value = UserForm_padron.dato_sello.Text
Cells(fila, 15).Value = UserForm_padron.dato_control_prenatal.Text
Cells(fila, 31).Value = UserForm_padron.dato_observaciones.Text

'para que el auditor pueda filtrar por A, B o C para completar el acta
If (UserForm_padron.dato_fuente.Text = "No consta fuente de información") Then
    Cells(fila, 14).Value = "A"
ElseIf (UserForm_padron.dato_fuente.Text = "Prestación inexistente") Then
    Cells(fila, 14).Value = "B"
ElseIf (UserForm_padron.dato_fuente.Text = "Caso duplicado") Then
    Cells(fila, 14).Value = "Caso duplicado"
ElseIf (UserForm_padron.dato_control_fuente.Text = "Fuente invalida") Then
    Cells(fila, 14).Value = "C"
ElseIf (UserForm_padron.dato_fuente.Text = "Caso duplicado") Then
    Cells(fila, 14).Value = "Caso duplicado"
ElseIf (UserForm_padron.dato_control_fuente.Text = "Fuente valida") Then
    Cells(fila, 14).Value = "Fuente valida"
Else
    Cells(fila, 14).Value = ""
End If

End Sub

'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_padron_verificacion_blancos() As Integer

If (UserForm_padron.dato_fuente.Text = "" Or UserForm_padron.dato_diagnostico.Text = "" Or UserForm_padron.dato_estudios.Text = "" Or UserForm_padron.dato_evaluacion_riesgo.Text = "" Or _
UserForm_padron.dato_ta.Text = "" Or UserForm_padron.dato_imc.Text = "" Or UserForm_padron.dato_percentilo.Text = "" Or UserForm_padron.dato_peso.Text = "" _
Or UserForm_padron.dato_talla.Text = "" Or UserForm_padron.dato_plan_seguimiento.Text = "" Or UserForm_padron.dato_tratamiento_instaurado.Text = "" Or UserForm_padron.dato_transcripcion.Text = "" _
Or UserForm_padron.dato_constancia_inmunizaciones.Text = "" Or UserForm_padron.dato_firma.Text = "" Or UserForm_padron.dato_sello.Text = "" Or UserForm_padron.dato_control_prenatal.Text = "") Then

userForm_padron_verificacion_blancos = 1

Else

userForm_padron_verificacion_blancos = 0

End If


End Function

'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userForm_padron_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

UserForm_padron.dato_diagnostico.Text = leyenda
UserForm_padron.dato_estudios.Text = leyenda
UserForm_padron.dato_evaluacion_riesgo.Text = leyenda
UserForm_padron.dato_ta.Text = leyenda
UserForm_padron.dato_imc.Text = leyenda
UserForm_padron.dato_percentilo.Text = leyenda
UserForm_padron.dato_peso.Text = leyenda
UserForm_padron.dato_talla.Text = leyenda
UserForm_padron.dato_plan_seguimiento.Text = leyenda
UserForm_padron.dato_tratamiento_instaurado.Text = leyenda
UserForm_padron.dato_transcripcion.Text = leyenda
UserForm_padron.dato_constancia_inmunizaciones.Text = leyenda
UserForm_padron.dato_firma.Text = leyenda
UserForm_padron.dato_sello.Text = leyenda
UserForm_padron.dato_control_prenatal.Text = leyenda

UserForm_padron.dato_diagnostico.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_estudios.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_evaluacion_riesgo.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_ta.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_imc.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_percentilo.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_peso.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_talla.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_plan_seguimiento.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_tratamiento_instaurado.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_transcripcion.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_constancia_inmunizaciones.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_firma.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_sello.BackColor = RGB(169, 169, 169)
UserForm_padron.dato_control_prenatal.BackColor = RGB(169, 169, 169)

UserForm_padron.dato_diagnostico.Locked = True
UserForm_padron.dato_estudios.Locked = True
UserForm_padron.dato_evaluacion_riesgo.Locked = True
UserForm_padron.dato_ta.Locked = True
UserForm_padron.dato_imc.Locked = True
UserForm_padron.dato_percentilo.Locked = True
UserForm_padron.dato_peso.Locked = True
UserForm_padron.dato_talla.Locked = True
UserForm_padron.dato_plan_seguimiento.Locked = True
UserForm_padron.dato_tratamiento_instaurado.Locked = True
UserForm_padron.dato_transcripcion.Locked = True
UserForm_padron.dato_constancia_inmunizaciones.Locked = True
UserForm_padron.dato_firma.Locked = True
UserForm_padron.dato_sello.Locked = True
UserForm_padron.dato_control_prenatal.Locked = True



End Function

'brief: protege y oculta celdas especificas
'param: void
'return: void

Sub proteger_y_ocultar()

Dim contraseña As String

contraseña = "hola"

Columns(13).Hidden = True
Columns(15).Hidden = True
Columns(17).Hidden = True
Columns(18).Hidden = True
Columns(19).Hidden = True
Columns(20).Hidden = True
Columns(21).Hidden = True
Columns(22).Hidden = True
Columns(23).Hidden = True
Columns(24).Hidden = True
Columns(25).Hidden = True
Columns(26).Hidden = True
Columns(27).Hidden = True
Columns(28).Hidden = True
Columns(29).Hidden = True
Columns(30).Hidden = True
Columns(31).Hidden = True
Columns(33).Hidden = True
Columns(34).Hidden = True
Columns(35).Hidden = True
Columns(36).Hidden = True
Columns(37).Hidden = True
Columns(38).Hidden = True
Columns(39).Hidden = True
Columns(40).Hidden = True
Columns(41).Hidden = True
Columns(42).Hidden = True
Columns(43).Hidden = True

ActiveSheet.Protect Password:=contraseña, DrawingObjects:=False, Contents:=True, Scenarios:= _
False, AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True


End Sub

'brief: desprotege y muestra celdas especificas
'param: void
'return: void

Sub desproteger_y_mostrar()

   
On Error Resume Next
    
    ActiveSheet.Unprotect Password:="hola"
    
    'MsgBox Err.Number
    
    If (Err.Number <> 1004) Then
    
        Columns(13).Hidden = False
        Columns(15).Hidden = False
        Columns(17).Hidden = False
        Columns(18).Hidden = False
        Columns(19).Hidden = False
        Columns(20).Hidden = False
        Columns(21).Hidden = False
        Columns(22).Hidden = False
        Columns(23).Hidden = False
        Columns(24).Hidden = False
        Columns(25).Hidden = False
        Columns(26).Hidden = False
        Columns(27).Hidden = False
        Columns(28).Hidden = False
        Columns(29).Hidden = False
        Columns(30).Hidden = False
        Columns(31).Hidden = False
        Columns(33).Hidden = False
        Columns(34).Hidden = False
        Columns(35).Hidden = False
        Columns(36).Hidden = False
        Columns(37).Hidden = False
        Columns(38).Hidden = False
        Columns(39).Hidden = False
        Columns(40).Hidden = False
        Columns(41).Hidden = False
        Columns(42).Hidden = False
        Columns(43).Hidden = False
        
    Else
    
    MsgBox ("Contraseña incorrecta")
    
    End If
    

End Sub

'brief: Analisa el formulario (viendo solo los motivos 1, 2 y 3). El motivo 4 debe ser verificado a mano
'param: void
'return: void


Sub analisis()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'11 es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ActiveSheet.Cells(6, 40).Value = Application.UserName
ActiveSheet.Cells(6, 41).Value = Date


'hace las siguientes lineas hasta que encuentra en la columna 12 (la del doble click) un celda vacia
Do Until ActiveSheet.Cells(i, 12).Value = ""
        
    'limpia las celdas de categoria y fundamento por si se utiliza la macro varias veces
    ActiveSheet.Cells(i, 39).Value = ""
    ActiveSheet.Cells(i, 41).Value = ""
    
    'los primeros 2 if son para los motivos 1, 2 y 3 respectivamente
    If (ActiveSheet.Cells(i, 14).Value = "A") Then
    
        ActiveSheet.Cells(i, 39).Value = 1
        ActiveSheet.Cells(i, 41).Value = "No consta fuente de información"
    
    ElseIf (ActiveSheet.Cells(i, 14).Value = "B") Then
    
        ActiveSheet.Cells(i, 39).Value = 2
        ActiveSheet.Cells(i, 41).Value = "Prestación inexistente"
        
    ElseIf (ActiveSheet.Cells(i, 14).Value = "C") Then
        
        ActiveSheet.Cells(i, 39).Value = 5
        ActiveSheet.Cells(i, 41).Value = "Fuente invalida"
        
    
    'para el motivo 4
    ElseIf (ActiveSheet.Cells(i, 12).Value = "Incompleto" Or ActiveSheet.Cells(i, 12).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
    
    
        'recorre la fila viendo que celdas estan vacias o dicen no y completa
        For j = 18 To 30
        
            If (ActiveSheet.Cells(i, j).Value = "" Or ActiveSheet.Cells(i, j).Value = "No") Then
                
                If (flag = 0) Then
                
                    leyenda = leyenda & ActiveSheet.Cells(10, j).Value
                    ActiveSheet.Cells(i, 39).Value = 3
                    flag = 1
                    
                Else
                
                    leyenda = leyenda & ", " & ActiveSheet.Cells(10, j).Value
                    
                End If
            
            End If
        
        Next
        
        
        If (flag = 1) Then
        
            ActiveSheet.Cells(i, 41).Value = leyenda
        End If
    
    
    'por si ocurre un caso duplicado
    ElseIf (ActiveSheet.Cells(i, 14).Value = "Caso duplicado") Then
        
        ActiveSheet.Cells(i, 39).Value = "Caso duplicado"
        ActiveSheet.Cells(i, 41).Value = "El caso debe ser eliminado de la muestra"
        
        
    Else
    
        MsgBox ("Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub


'brief: desbloquea los campos con datos obligatorios, les pone un fondo blanco y hace que valgan ""
'param: el codigo declarado donde se hizo doble click
'return: void

Sub userForm_padron_permitir_campos_requeridos(ByVal codigo As String)

Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim cantidad_codigos_normales As Integer
Dim cantidad_requerimientos As Integer
Dim fila_codigos_especiales_inicio As Integer
Dim fila_codigos_especiales_final As Integer
Dim texto As String
Dim leyenda As String
Dim leyenda2 As String
Dim leyenda3 As String
Dim leyenda4 As String
Dim leyenda5 As String
Dim parte_codigo As String
Dim flag As Boolean


cantidad_codigos_normales = 212
cantidad_requerimientos = 31
fila_codigos_especiales_inicio = 214
fila_codigos_especiales_final = 216
flag = False

leyenda = "Labrar acta"
leyenda2 = "Labrar acta e indicar fuente de información en observaciones"
leyenda3 = "Caso duplicado"
leyenda4 = "Dato no obligatorio"
leyenda5 = "Fuente invalida"

parte_codigo = Left(codigo, 3)


x = 1


'para limpiar el textbox de diagnostico cuando se cambia de "no consta fuente de informacion" o "prestacion inexistente"
'a una fuente de informacion
If (UserForm_padron.dato_validacion.Text <> leyenda And UserForm_padron.dato_validacion.Text <> leyenda2 And _
UserForm_padron.dato_validacion.Text <> leyenda3) Then

    If (UserForm_padron.dato_diagnostico.Text = leyenda4) Then
    
    UserForm_padron.dato_diagnostico.Locked = False
    UserForm_padron.dato_diagnostico.BackColor = RGB(255, 255, 255)
    UserForm_padron.dato_diagnostico.Text = ""
    
    End If
    
End If

'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (UserForm_padron.dato_validacion.Text <> leyenda And UserForm_padron.dato_validacion.Text <> leyenda2 And UserForm_padron.dato_validacion.Text <> leyenda3) Then

    For i = 1 To cantidad_codigos_normales
    
        'coincidencia de poblacion con el codigo
        If ((ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = ActiveSheet.Cells(filaDobleClick, columnaDobleClick + cantidad_requerimientos).Value And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo)) Then
                
            For j = 5 To cantidad_requerimientos
                'verifica que que el dato que corresponde a la celda se obligatorio
                If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                    
                    'copia el nombre del dato obligatorio y empiesa a verificar cual es
                    texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                    
                    If (texto = "Peso") Then
                    
                        With UserForm_padron.dato_peso
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                        .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                
                    
                    If (texto = "Talla") Then
                    
                        With UserForm_padron.dato_talla
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                        .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "TA") Then
                    
                        With UserForm_padron.dato_ta
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "IMC") Then
                    
                        With UserForm_padron.dato_imc
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Informe o transcripción de estudios solicitados") Then
                    
                        With UserForm_padron.dato_transcripcion
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Evaluacion de riesgo") Then
                    
                        With UserForm_padron.dato_evaluacion_riesgo
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Diagnostico") Then
                    
                        With UserForm_padron.dato_diagnostico
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Tratamiento instaurado") Then
                    
                        With UserForm_padron.dato_tratamiento_instaurado
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Plan de seguimiento") Then
                    
                        With UserForm_padron.dato_plan_seguimiento
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Constancia de aplicación de inmunizaciones") Then
                    
                        With UserForm_padron.dato_constancia_inmunizaciones
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Firma") Then
                    
                        With UserForm_padron.dato_firma
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Sello") Then
                    
                        With UserForm_padron.dato_sello
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
'                    If (texto = "Altura uterina") Then
'
'                        With UserForm_padron
'                        .Locked = False
'                        If (.Text = "Dato no obligatorio") Then
'                            .Text = ""
'                        End If
'                        .BackColor = RGB(255, 255, 255)
'                        End With
'
'                    End If
'
'                    If (texto = "Apgar") Then
'
'                        With UserForm_padron
'                        .Locked = False
'                        If (.Text = "Dato no obligatorio") Then
'                            .Text = ""
'                        End If
'                        .BackColor = RGB(255, 255, 255)
'                        End With
'
'                    End If
'
'                    If (texto = "Calculo de amenorrea") Then
'
'                        With UserForm_padron.dato_amen
'                        .Locked = False
'                        If (.Text = "Dato no obligatorio") Then
'                            .Text = ""
'                        End If
'                        .BackColor = RGB(255, 255, 255)
'                        End With
'
'                    End If

                    If (texto = "Constancia de solicitud de screening neonatal") Then

                        With UserForm_padron.dato_peso
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With

                    End If

                    If (texto = "Diagnostico de embarazo de alto riesgo") Then

                        With UserForm_padron.dato_peso
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With

                    End If

'                    If (texto = "Perimetro cefalico") Then
'
'                        With UserForm_padron.dato_p
'                        .Locked = False
'                        If (.Text = "Dato no obligatorio") Then
'                            .Text = ""
'                        End If
'                        .BackColor = RGB(255, 255, 255)
'                        End With
'
'                    End If
                    
                    If (texto = "Percentilo") Then
                    
                        With UserForm_padron.dato_percentilo
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "N de Control Prenatal") Then
                        
                        With UserForm_padron.dato_control_prenatal
                        .Locked = False
                        If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                        End If
                        .BackColor = RGB(255, 255, 255)
                        End With
                        
                    End If
                    
                    If (texto = "Examen mamario" Or texto = "Evaluacion genitourinaria" _
                    Or texto = "Odontograma" Or texto = "Medicion de agudeza visual" _
                    Or texto = "Colonoscopia") Then
                            
                        With UserForm_padron.dato_estudios
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            
                            If (estudiosFlag = 0) Then
                                .Clear
                                Select Case (texto)

                                    Case "Examen mamario"
                                        .AddItem "Examen mamario"
                                        .AddItem "Evaluacion genitourinaria"
                                        .AddItem "Examen mamario y evaluacion genitourinaria"
                                        
                                    Case "Evaluacion genitourinaria"
                                    .AddItem "Evaluacion genitourinaria"
                                    
                                    Case "Medicion de agudeza visual"
                                            .AddItem "Medicion de agudeza visual"
                                    
                                    Case "Odontograma"
                                        .AddItem "Odontograma"
                                        
                                    Case "Colonoscopia"
                                        .AddItem "Colonoscopia"
                                        
                                End Select
                                
                                .AddItem "No consta"
                                
                                estudiosFlag = 1
                            End If

                        End With
                            
                    End If
                    
                End If
                
            Next
            
            flag = True
            'para romper el primer for (el de i)
            i = cantidad_codigos_normales
            
        End If
        
    Next
                
    If (flag = False And (parte_codigo = "PRP" Or parte_codigo = "IGR" Or parte_codigo = "LBL")) Then
            
        For i = fila_codigos_especiales_inicio To fila_codigos_especiales_final

            'coincidencia de poblacion con el codigo
            If (ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = parte_codigo) Then
                    
                For j = 5 To cantidad_requerimientos
                    'verifica que que el dato que corresponde a la celda se obligatorio
                    If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                        
                        'copia el nombre del dato obligatorio y empiesa a verificar cual es
                        texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                        
                        If (texto = "Peso") Then
                        
                            With UserForm_padron.dato_peso
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                    
                        
                        If (texto = "Talla") Then
                        
                            With UserForm_padron.dato_talla
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                            .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "TA") Then
                        
                            With UserForm_padron.dato_ta
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "IMC") Then
                        
                            With UserForm_padron.dato_imc
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Informe o transcripción de estudios solicitados") Then
                        
                            With UserForm_padron.dato_transcripcion
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Evaluacion de riesgo") Then
                        
                            With UserForm_padron.dato_evaluacion_riesgo
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Diagnostico") Then
                        
                            With UserForm_padron.dato_diagnostico
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Tratamiento instaurado") Then
                        
                            With UserForm_padron.dato_tratamiento_instaurado
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Plan de seguimiento") Then
                        
                            With UserForm_padron.dato_plan_seguimiento
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Constancia de aplicación de inmunizaciones") Then
                        
                            With UserForm_padron.dato_constancia_inmunizaciones
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Firma") Then
                        
                            With UserForm_padron.dato_firma
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Sello") Then
                        
                            With UserForm_padron.dato_sello
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
    '                    If (texto = "Altura uterina") Then
    '
    '                        With UserForm_padron
    '                        .Locked = False
    '                        If (.Text = "Dato no obligatorio") Then
    '                            .Text = ""
    '                        End If
    '                        .BackColor = RGB(255, 255, 255)
    '                        End With
    '
    '                    End If
    '
    '                    If (texto = "Apgar") Then
    '
    '                        With UserForm_padron
    '                        .Locked = False
    '                        If (.Text = "Dato no obligatorio") Then
    '                            .Text = ""
    '                        End If
    '                        .BackColor = RGB(255, 255, 255)
    '                        End With
    '
    '                    End If
    '
    '                    If (texto = "Calculo de amenorrea") Then
    '
    '                        With UserForm_padron.dato_amen
    '                        .Locked = False
    '                        If (.Text = "Dato no obligatorio") Then
    '                            .Text = ""
    '                        End If
    '                        .BackColor = RGB(255, 255, 255)
    '                        End With
    '
    '                    End If
    
                        If (texto = "Constancia de solicitud de screening neonatal") Then
    
                            With UserForm_padron.dato_peso
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
    
                        End If
    
                        If (texto = "Diagnostico de embarazo de alto riesgo") Then
    
                            With UserForm_padron.dato_peso
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
    
                        End If
    
    '                    If (texto = "Perimetro cefalico") Then
    '
    '                        With UserForm_padron.dato_p
    '                        .Locked = False
    '                        If (.Text = "Dato no obligatorio") Then
    '                            .Text = ""
    '                        End If
    '                        .BackColor = RGB(255, 255, 255)
    '                        End With
    '
    '                    End If
                        
                        If (texto = "Percentilo") Then
                        
                            With UserForm_padron.dato_percentilo
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "N de Control Prenatal") Then
                            
                            With UserForm_padron.dato_control_prenatal
                            .Locked = False
                            If (.Text = "Dato no obligatorio") Then
                                .Text = ""
                            End If
                            .BackColor = RGB(255, 255, 255)
                            End With
                            
                        End If
                        
                        If (texto = "Examen mamario" Or texto = "Evaluacion genitourinaria" _
                        Or texto = "Odontograma" Or texto = "Medicion de agudeza visual" _
                        Or texto = "Colonoscopia") Then
                                
                            With UserForm_padron.dato_estudios
                                .Locked = False
                                If (.Text = "Dato no obligatorio") Then
                                    .Text = ""
                                End If
                                .BackColor = RGB(255, 255, 255)
                                
                                If (estudiosFlag = 0) Then
                                    .Clear
                                    Select Case (texto)
    
                                        Case "Examen mamario"
                                            .AddItem "Examen mamario"
                                            .AddItem "Evaluacion genitourinaria"
                                            .AddItem "Examen mamario y evaluacion genitourinaria"
                                            
                                        Case "Evaluacion genitourinaria"
                                        .AddItem "Evaluacion genitourinaria"
                                        
                                        Case "Medicion de agudeza visual"
                                                .AddItem "Medicion de agudeza visual"
                                        
                                        Case "Odontograma"
                                            .AddItem "Odontograma"
                                            
                                        Case "Colonoscopia"
                                            .AddItem "Colonoscopia"
                                            
                                    End Select
                                    
                                    .AddItem "No consta"
                                    
                                    estudiosFlag = 1
                                End If
    
                            End With
                                
                        End If
                        
                    End If
                
                Next
        
                flag = True
                'para romper el primer for (el de i)
                i = fila_codigos_especiales_final
            
            End If
    
        Next

    End If

End If

End Sub


