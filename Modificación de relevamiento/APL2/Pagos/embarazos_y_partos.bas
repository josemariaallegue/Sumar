Attribute VB_Name = "embarazos_y_partos"
'declaracion de variables globales

Public filaDobleClickEYP As Integer

Public columnaDobleClickEYP As Integer

Public codigoDobleClickEYP As String

Public estudiosFlagEYP  As Integer

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteEYP As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorEYP As Integer


'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub copiar_eyp_datos_fijos(ByVal Target As Range)

With userForm_eyp

    .TextBox_n_efector.Text = Cells(Target.Row, Target.Column - 11).Value
    .TextBox_denominacion_efector.Text = Cells(Target.Row, Target.Column - 10).Value
    .TextBox_orden_pago_factura = Cells(Target.Row, Target.Column - 9).Value
    .TextBox_beneficiario.Text = Cells(Target.Row, Target.Column - 8).Value & " " & Cells(Target.Row, Target.Column - 7).Value
    .TextBox_documento.Text = Cells(Target.Row, Target.Column - 6).Value
    .TextBox_clave_beneficiario.Text = Cells(Target.Row, Target.Column - 5).Value
    .TextBox_codigo.Text = Cells(Target.Row, Target.Column - 3).Value
    .TextBox_descripcion.Text = Cells(Target.Row, Target.Column - 2).Value
    .TextBox_fecha_prestacion.Text = Cells(Target.Row, Target.Column - 1).Value
    .TextBox_fecha_nacimiento.Text = Cells(Target.Row, Target.Column + 3).Value
    .TextBox_monto = Cells(Target.Row, Target.Column + 25).Value
    .TextBox_edad.Text = Cells(Target.Row, Target.Column + 27).Value

End With

'verifica que si la celda de nombre de prestacion contiene "La prestacion no corresponde al grupo poblacional" si
'esto es verdadero le otorga un valor al textbox de control de fuente de informacion
If (Cells(filaDobleClickEYP, columnaDobleClickEYP - 2).Value = "La prestación no corresponde al grupo poblacional") Then

    userForm_eyp.dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"
    
End If

End Sub


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_eyp_bloquear()

With userForm_eyp

    .TextBox_n_efector.Locked = True
    .TextBox_denominacion_efector.Locked = True
    .TextBox_orden_pago_factura.Locked = True
    .TextBox_beneficiario.Locked = True
    .TextBox_clave_beneficiario.Locked = True
    .TextBox_documento.Locked = True
    .TextBox_codigo.Locked = True
    .TextBox_fecha_prestacion.Locked = True
    .TextBox_fecha_nacimiento.Locked = True
    .TextBox_edad.Locked = True
    .TextBox_monto.Locked = True
    .TextBox_descripcion.Locked = True
    
    '.dato_fuente.Locked = True
    '.dato_diagnostico.Locked = True
    
    .dato_n_control.Locked = True
    .dato_peso.Locked = True
    .dato_talla.Locked = True
    .dato_ta.Locked = True
    .dato_imc.Locked = True
    .dato_apgar.Locked = True
    .dato_perimetro_cefalico.Locked = True
    .dato_altura_uterina.Locked = True
    .dato_transcripcion_estudios.Locked = True
    .dato_amenorrea.Locked = True
    .dato_estudios.Locked = True
    .dato_evaluacion_riesgo.Locked = True
    .dato_vida_fetal.Locked = True
    .dato_tratamiento_instaurado.Locked = True
    .dato_plan_seguimiento.Locked = True
    .dato_screening_neonatal.Locked = True
    .dato_constancia_inmunizaciones.Locked = True
    .dato_fecha_notificacion.Locked = True
    .dato_firma.Locked = True
    .dato_sello.Locked = True

End With

End Function

'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_eyp_copiar_datos_relevamiento(ByVal Target As Range)


Dim leyenda As String

leyenda = "Dato no obligatorio"


userForm_eyp.dato_fuente.Text = Cells(Target.Row, Target.Column + 1).Value

'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
If (userForm_eyp.dato_fuente.Text <> "No consta fuente de información" And userForm_eyp.dato_fuente.Text <> "Prestación inexistente" And _
userForm_eyp.dato_control_fuente.Text <> "Fuente invalida") Then
    
    userForm_eyp.dato_diagnostico.Text = Cells(Target.Row, Target.Column + 16).Value
    
    'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
    'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
    'lineas de codigo de mas
    userForm_eyp.dato_n_control.Text = Cells(Target.Row, Target.Column + 4).Value
    If (userForm_eyp.dato_n_control.Text = "") Then
        userForm_eyp.dato_n_control.Text = leyenda
    End If
    
    userForm_eyp.dato_peso.Text = Cells(Target.Row, Target.Column + 5).Value
    If (userForm_eyp.dato_peso.Text = "") Then
        userForm_eyp.dato_peso.Text = leyenda
    End If
    
    userForm_eyp.dato_talla.Text = Cells(Target.Row, Target.Column + 6).Value
    If (userForm_eyp.dato_talla.Text = "") Then
        userForm_eyp.dato_talla.Text = leyenda
    End If
    
    userForm_eyp.dato_ta.Text = Cells(Target.Row, Target.Column + 7).Value
    If (userForm_eyp.dato_ta.Text = "") Then
        userForm_eyp.dato_ta.Text = leyenda
    End If
    
    userForm_eyp.dato_imc.Text = Cells(Target.Row, Target.Column + 8).Value
    If (userForm_eyp.dato_imc.Text = "") Then
        userForm_eyp.dato_imc.Text = leyenda
    End If
    
    userForm_eyp.dato_apgar.Text = Cells(Target.Row, Target.Column + 9).Value
    If (userForm_eyp.dato_apgar.Text = "") Then
        userForm_eyp.dato_apgar.Text = leyenda
    End If
    
    userForm_eyp.dato_perimetro_cefalico.Text = Cells(Target.Row, Target.Column + 10).Value
    If (userForm_eyp.dato_perimetro_cefalico.Text = "") Then
        userForm_eyp.dato_perimetro_cefalico.Text = leyenda
    End If
    
    userForm_eyp.dato_altura_uterina.Text = Cells(Target.Row, Target.Column + 11).Value
    If (userForm_eyp.dato_altura_uterina.Text = "") Then
        userForm_eyp.dato_altura_uterina.Text = leyenda
    End If
    
    userForm_eyp.dato_transcripcion_estudios.Text = Cells(Target.Row, Target.Column + 12).Value
    If (userForm_eyp.dato_transcripcion_estudios.Text = "") Then
        userForm_eyp.dato_transcripcion_estudios.Text = leyenda
    End If
    
    userForm_eyp.dato_amenorrea.Text = Cells(Target.Row, Target.Column + 13).Value
    If (userForm_eyp.dato_amenorrea.Text = "") Then
        userForm_eyp.dato_amenorrea.Text = leyenda
    End If
    
    userForm_eyp.dato_estudios.Text = Cells(Target.Row, Target.Column + 14).Value
    If (userForm_eyp.dato_estudios.Text = "") Then
        userForm_eyp.dato_estudios.Text = leyenda
    End If
    
    userForm_eyp.dato_evaluacion_riesgo.Text = Cells(Target.Row, Target.Column + 15).Value
    If (userForm_eyp.dato_evaluacion_riesgo.Text = "") Then
        userForm_eyp.dato_evaluacion_riesgo.Text = leyenda
    End If
    
    userForm_eyp.dato_vida_fetal.Text = Cells(Target.Row, Target.Column + 17).Value
    If (userForm_eyp.dato_vida_fetal.Text = "") Then
        userForm_eyp.dato_vida_fetal.Text = leyenda
    End If
    
    userForm_eyp.dato_tratamiento_instaurado.Text = Cells(Target.Row, Target.Column + 18).Value
    If (userForm_eyp.dato_tratamiento_instaurado.Text = "") Then
        userForm_eyp.dato_tratamiento_instaurado.Text = leyenda
    End If
    
    userForm_eyp.dato_plan_seguimiento.Text = Cells(Target.Row, Target.Column + 19).Value
    If (userForm_eyp.dato_plan_seguimiento.Text = "") Then
        userForm_eyp.dato_plan_seguimiento.Text = leyenda
    End If
    
    userForm_eyp.dato_screening_neonatal.Text = Cells(Target.Row, Target.Column + 20).Value
    If (userForm_eyp.dato_screening_neonatal.Text = "") Then
        userForm_eyp.dato_screening_neonatal.Text = leyenda
    End If
    
    userForm_eyp.dato_constancia_inmunizaciones.Text = Cells(Target.Row, Target.Column + 21).Value
    If (userForm_eyp.dato_constancia_inmunizaciones.Text = "") Then
        userForm_eyp.dato_constancia_inmunizaciones.Text = leyenda
    End If
    
    userForm_eyp.dato_fecha_notificacion.Text = Cells(Target.Row, Target.Column + 22).Value
    If (userForm_eyp.dato_fecha_notificacion.Text = "") Then
        userForm_eyp.dato_fecha_notificacion.Text = leyenda
    End If
    
    userForm_eyp.dato_firma.Text = Cells(Target.Row, Target.Column + 23).Value
    If (userForm_eyp.dato_firma.Text = "") Then
        userForm_eyp.dato_firma.Text = leyenda
    End If
    
    userForm_eyp.dato_sello.Text = Cells(Target.Row, Target.Column + 24).Value
    If (userForm_eyp.dato_sello.Text = "") Then
        userForm_eyp.dato_sello.Text = leyenda
    End If
    
    'userForm_eyp.dato_control_fuente.Text = Cells(target.Row, target.Column + 2).Value
    
    'userForm_eyp.dato_validacion.Text = Cells(target.Row, target.Column + 21).Value
    
Else
    
    Call userform_eyp_dato_no_obligatorio
    
End If

userForm_eyp.dato_observaciones.Text = Cells(Target.Row, Target.Column + 26).Value

End Sub

'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_eyp_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_eyp

    .dato_n_control.Text = leyenda
    .dato_peso.Text = leyenda
    .dato_talla.Text = leyenda
    .dato_ta.Text = leyenda
    .dato_imc.Text = leyenda
    .dato_apgar.Text = leyenda
    .dato_perimetro_cefalico.Text = leyenda
    .dato_altura_uterina.Text = leyenda
    .dato_transcripcion_estudios.Text = leyenda
    .dato_amenorrea.Text = leyenda
    .dato_estudios.Text = leyenda
    .dato_evaluacion_riesgo.Text = leyenda
    .dato_diagnostico.Text = leyenda
    .dato_vida_fetal.Text = leyenda
    .dato_tratamiento_instaurado.Text = leyenda
    .dato_plan_seguimiento.Text = leyenda
    .dato_screening_neonatal.Text = leyenda
    .dato_constancia_inmunizaciones.Text = leyenda
    .dato_fecha_notificacion.Text = leyenda
    .dato_firma.Text = leyenda
    .dato_sello.Text = leyenda
    
    .dato_n_control.BackColor = RGB(169, 169, 169)
    .dato_peso.BackColor = RGB(169, 169, 169)
    .dato_talla.BackColor = RGB(169, 169, 169)
    .dato_ta.BackColor = RGB(169, 169, 169)
    .dato_imc.BackColor = RGB(169, 169, 169)
    .dato_apgar.BackColor = RGB(169, 169, 169)
    .dato_perimetro_cefalico.BackColor = RGB(169, 169, 169)
    .dato_altura_uterina.BackColor = RGB(169, 169, 169)
    .dato_transcripcion_estudios.BackColor = RGB(169, 169, 169)
    .dato_amenorrea.BackColor = RGB(169, 169, 169)
    .dato_estudios.BackColor = RGB(169, 169, 169)
    .dato_evaluacion_riesgo.BackColor = RGB(169, 169, 169)
    .dato_diagnostico.BackColor = RGB(169, 169, 169)
    .dato_vida_fetal.BackColor = RGB(169, 169, 169)
    .dato_tratamiento_instaurado.BackColor = RGB(169, 169, 169)
    .dato_plan_seguimiento.BackColor = RGB(169, 169, 169)
    .dato_screening_neonatal.BackColor = RGB(169, 169, 169)
    .dato_constancia_inmunizaciones.BackColor = RGB(169, 169, 169)
    .dato_fecha_notificacion.BackColor = RGB(169, 169, 169)
    .dato_firma.BackColor = RGB(169, 169, 169)
    .dato_sello.BackColor = RGB(169, 169, 169)
    
    .dato_n_control.Locked = True
    .dato_peso.Locked = True
    .dato_talla.Locked = True
    .dato_ta.Locked = True
    .dato_imc.Locked = True
    .dato_apgar.Locked = True
    .dato_perimetro_cefalico.Locked = True
    .dato_altura_uterina.Locked = True
    .dato_transcripcion_estudios.Locked = True
    .dato_amenorrea.Locked = True
    .dato_estudios.Locked = True
    .dato_evaluacion_riesgo.Locked = True
    .dato_diagnostico.Locked = True
    .dato_vida_fetal.Locked = True
    .dato_tratamiento_instaurado.Locked = True
    .dato_plan_seguimiento.Locked = True
    .dato_screening_neonatal.Locked = True
    .dato_constancia_inmunizaciones.Locked = True
    .dato_fecha_notificacion.Locked = True
    .dato_firma.Locked = True
    .dato_sello.Locked = True
    
End With

End Function


'brief: copia los datos del userform_eyp al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_eyp_guardar_datos(ByVal fila As Integer)

With userForm_eyp
    
    Cells(fila, 14).Value = .dato_fuente.Text
    Cells(fila, 17).Value = .dato_n_control.Text
    Cells(fila, 18).Value = .dato_peso.Text
    Cells(fila, 19).Value = .dato_talla.Text
    Cells(fila, 20).Value = .dato_ta.Text
    Cells(fila, 21).Value = .dato_imc.Text
    Cells(fila, 22).Value = .dato_apgar.Text
    Cells(fila, 23).Value = .dato_perimetro_cefalico.Text
    Cells(fila, 24).Value = .dato_altura_uterina.Text
    Cells(fila, 25).Value = .dato_transcripcion_estudios.Text
    Cells(fila, 26).Value = .dato_amenorrea.Text
    Cells(fila, 27).Value = .dato_estudios.Text
    Cells(fila, 28).Value = .dato_evaluacion_riesgo.Text
    Cells(fila, 29).Value = .dato_diagnostico.Text
    Cells(fila, 30).Value = .dato_vida_fetal.Text
    Cells(fila, 31).Value = .dato_tratamiento_instaurado.Text
    Cells(fila, 32).Value = .dato_plan_seguimiento.Text
    Cells(fila, 33).Value = .dato_screening_neonatal.Text
    Cells(fila, 34).Value = .dato_constancia_inmunizaciones.Text
    Cells(fila, 35).Value = .dato_fecha_notificacion.Text
    Cells(fila, 36).Value = .dato_firma.Text
    Cells(fila, 37).Value = .dato_sello.Text
    Cells(fila, 39).Value = .dato_observaciones.Text
    
    'para que el auditor pueda filtrar por A, B o C para completar el acta
    If (.dato_fuente.Text = "No consta fuente de información") Then
        Cells(fila, 15).Value = "A"
        
    ElseIf (.dato_fuente.Text = "Prestación inexistente") Then
        Cells(fila, 15).Value = "B"
        
    ElseIf (.dato_fuente.Text = "Caso duplicado") Then
        Cells(fila, 15).Value = "Caso duplicado"
        
    ElseIf (.dato_control_fuente.Text = "Fuente invalida") Then
        Cells(fila, 15).Value = "C"
        
    ElseIf (.dato_control_fuente.Text = "Fuente valida") Then
        Cells(fila, 15).Value = "Fuente valida"
        
    Else
        Cells(fila, 15).Value = ""
    End If
    
End With

End Sub


'brief desbloquea y limpia los combobox y textbos que corresponde a los datos obligatorios
'param void
'return void

Sub userForm_eyp_permitir_campos_requeridos(ByVal codigo As String)

Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim texto As String
Dim leyenda As String
Dim leyenda2 As String
Dim leyenda3 As String
Dim leyenda4 As String

leyenda = "Labrar acta"
leyenda2 = "Labrar acta e indicar fuente de información en observaciones"
leyenda3 = "Fuente invalida"
leyenda4 = "Dato no obligatorio"


x = 1

'para limpiar el textbox de diagnostico cuando se cambia de "no consta fuente de informacion" o "prestacion inexistente"
'a una fuente de informacion
If (userForm_eyp.dato_validacion.Text <> leyenda And userForm_eyp.dato_validacion.Text <> leyenda2 And _
userForm_eyp.dato_validacion.Text <> leyenda3) Then

    If (userForm_eyp.dato_diagnostico.Text = leyenda4) Then
    
        userForm_eyp.dato_diagnostico.Locked = False
        userForm_eyp.dato_diagnostico.BackColor = RGB(255, 255, 255)
        userForm_eyp.dato_diagnostico.Text = ""
    
    End If
    
End If


'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_eyp.dato_validacion.Text <> leyenda And userForm_eyp.dato_validacion.Text <> leyenda2 And userForm_eyp.dato_validacion.Text <> leyenda3) Then

    For i = 1 To 250
    
        'coincidencia de poblacion con el codigo o el codigo con la poblacion embarazo
        If ((ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = ActiveSheet.Cells(filaDobleClickEYP, columnaDobleClickEYP + 31).Value And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo) Or (ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = "Embarazos")) Then
                
                've si el codigo es CTC001A97 y la poblacion es niños (distintos datos obligatorios por edad)
                If (codigo = "CTC001A97" And ActiveSheet.Cells(filaDobleClickEYP, columnaDobleClickEYP + 31).Value = "Niños") Then
                    
                    'hace coincidir la edad en el relevamiento con la descripcion
                    If ((ActiveSheet.Cells(filaDobleClickEYP, columnaDobleClickEYP + 27).Value < 1 And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "menores de 1") Or _
                    ((ActiveSheet.Cells(filaDobleClickEYP, columnaDobleClickEYP + 27).Value >= 1 And ActiveSheet.Cells(filaDobleClickEYP, columnaDobleClickEYP + 27).Value < 6) And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "1 a 5") Or _
                    (ActiveSheet.Cells(filaDobleClickEYP, columnaDobleClickEYP + 27).Value >= 6 And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "mayores a 6")) Then

                        For j = 5 To 32
                            
                            'verifica que que el dato que corresponde a la celda se obligatorio
                            If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                                
                                'copia el nombre del dato obligatorio y empiesa a verificar cual es
                                texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                                
                                Select Case texto
                                
                                    Case "N de Control Prenatal"

                                        With userForm_eyp.dato_n_control
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                
                                    Case "Peso"

                                    
                                        With userForm_eyp.dato_peso
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Talla"
                                    
                                        With userForm_eyp.dato_talla
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "TA"
                                        
                                        With userForm_eyp.dato_ta
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "IMC"
                                    
                                        With userForm_eyp.dato_imc
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Apgar"

                                        With userForm_eyp.dato_apgar
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    
                                    Case "Perimetro cefalico"

                                        With userForm_eyp.dato_perimetro_cefalico
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Altura uterina"

                                        With userForm_eyp.dato_altura_uterina
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Informe o transcripción de estudios solicitados"
                                        
                                        With userForm_eyp.dato_transcripcion_estudios
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Calculo de amenorrea"

                                        With userForm_eyp.dato_amenorrea
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Examen mamario", "Evaluacion genitourinaria", "Odontograma", "Medicion de agudeza visual"
                                        
                                        With userForm_eyp.dato_estudios
                                                .Locked = False
                                                If (.Text = "Dato no obligatorio") Then
                                                    .Text = ""
                                                End If
                                                .BackColor = RGB(255, 255, 255)
                                            
                                            If (estudiosFlagEYP = 0) Then
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
                                                
                                                estudiosFlagEYP = 1
                                            End If
    
                                        End With
                                    
                                    
                                    Case "Evaluacion de riesgo"
                                    
                                        With userForm_eyp.dato_evaluacion_riesgo
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
'                                    Case "Diagnostico"
'
'                                        With userForm_eyp.dato_diagnostico
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                        End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
                                        
                                        
                                    Case "Diagnostico de vida fetal"

                                        With userForm_eyp.dato_peso
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Tratamiento instaurado"
                                        
                                        With userForm_eyp.dato_tratamiento_instaurado
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Plan de seguimiento"
                                    
                                        With userForm_eyp.dato_plan_seguimiento
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Constancia de solicitud de screening neonatal"

                                        With userForm_eyp.dato_screening_neonatal
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                    
                                    
                                    Case "Constancia de aplicación de inmunizaciones"
                                    
                                        With userForm_eyp.dato_constancia_inmunizaciones
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    Case "Fecha de notificación o de parto/cesárea"
                                        
                                        With userForm_eyp.dato_fecha_notificacion
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    Case "Firma"
                                        
                                        With userForm_eyp.dato_firma
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    
                                    Case "Sello"
                                        
                                        With userForm_eyp.dato_sello
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                        
'                                    Case "Percentilo"
'
'                                        With userForm_eyp.dato_percentilo
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                        End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
                                        
                                    
                                End Select
                                
                            End If

                        Next
                        
                        'para romper el primer for (el de i)
                        i = 250
                    
                    End If
                    
                
                
                Else

                
                    For j = 5 To 32
                        
                        'verifica que que el dato que corresponde a la celda se obligatorio
                        If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                            
                            'copia el nombre del dato obligatorio y empiesa a verificar cual es
                            texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                            
                            Select Case texto
                                
                                Case "N de Control Prenatal"

                                        With userForm_eyp.dato_n_control
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                
                                    Case "Peso"
       
                                        With userForm_eyp.dato_peso
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Talla"
                                        
                                        With userForm_eyp.dato_talla
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "TA"
                                        
                                        With userForm_eyp.dato_ta
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "IMC"
                                    
                                        With userForm_eyp.dato_imc
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Apgar"

                                        With userForm_eyp.dato_apgar
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    
                                    Case "Perimetro cefalico"

                                        With userForm_eyp.dato_perimetro_cefalico
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Altura uterina"

                                        With userForm_eyp.dato_altura_uterina
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Informe o transcripción de estudios solicitados"
                                        
                                        With userForm_eyp.dato_transcripcion_estudios
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Calculo de amenorrea"

                                        With userForm_eyp.dato_amenorrea
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Examen mamario", "Evaluacion genitourinaria", "Odontograma", "Medicion de agudeza visual"
                                        
                                        With userForm_eyp.dato_estudios
                                                .Locked = False
                                                If (.Text = "Dato no obligatorio") Then
                                                    .Text = ""
                                                End If
                                                .BackColor = RGB(255, 255, 255)
                                            
                                            If (estudiosFlagEYP = 0) Then
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
                                                
                                                estudiosFlagEYP = 1
                                            End If
    
                                        End With
                                    
                                    
                                    Case "Evaluacion de riesgo"
                                    
                                        With userForm_eyp.dato_evaluacion_riesgo
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
'                                    Case "Diagnostico"
'
'                                        With userForm_eyp.dato_diagnostico
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                        End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
                                        
                                        
                                    Case "Diagnostico de vida fetal"

                                        With userForm_eyp.dato_vida_fetal
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Tratamiento instaurado"
                                        
                                        With userForm_eyp.dato_tratamiento_instaurado
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Plan de seguimiento"
                                    
                                        With userForm_eyp.dato_plan_seguimiento
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                        
                                    Case "Constancia de solicitud de screening neonatal"

                                        With userForm_eyp.dato_screening_neonatal
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                         End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                    
                                    
                                    Case "Constancia de aplicación de inmunizaciones"
                                    
                                        With userForm_eyp.dato_constancia_inmunizaciones
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                    
                                    Case "Fecha de notificación o de parto/cesárea"
                                        
                                        With userForm_eyp.dato_fecha_notificacion
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    Case "Firma"
                                        
                                        With userForm_eyp.dato_firma
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    
                                    Case "Sello"
                                        
                                        With userForm_eyp.dato_sello
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                        
'                                    Case "Percentilo"
'
'                                        With userForm_eyp.dato_percentilo
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                        End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
                                    
                            End Select
                            
                        End If
                            
                    Next
                    
                    'para romper el primer for (el de i)
                    i = 250
                    
                End If
        
        End If
        
    Next
       
End If

End Sub


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_eyp_verificacion_blancos() As Integer

With userForm_eyp

    If (.dato_fuente.Text = "" Or .dato_diagnostico.Text = "" Or .dato_estudios.Text = "" Or .dato_n_control.Text = "" _
    Or .dato_peso.Text = "" Or .dato_talla.Text = "" Or .dato_ta.Text = "" Or .dato_imc.Text = "" Or .dato_apgar.Text = "" _
    Or .dato_perimetro_cefalico.Text = "" Or .dato_altura_uterina.Text = "" Or .dato_transcripcion_estudios.Text = "" _
    Or .dato_amenorrea.Text = "" Or .dato_evaluacion_riesgo.Text = "" Or .dato_vida_fetal.Text = "" Or .dato_tratamiento_instaurado.Text = "" _
    Or .dato_plan_seguimiento.Text = "" Or .dato_screening_neonatal.Text = "" Or .dato_constancia_inmunizaciones.Text = "" _
    Or .dato_fecha_notificacion.Text = "" Or .dato_firma.Text = "" Or .dato_sello.Text = "") Then
        
        userForm_eyp_verificacion_blancos = 1

    Else

        userForm_eyp_verificacion_blancos = 0
    
    End If
 
End With

End Function
