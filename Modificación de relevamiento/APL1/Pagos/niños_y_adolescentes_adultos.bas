Attribute VB_Name = "niños_y_adolescentes_adultos"
'declaracion de variables globales

Public filaDobleClick As Integer

Public columnaDobleClick As Integer

Public codigoDobleClick As String

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistente As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public error As Integer


'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub copiar_naa_datos_fijos(ByVal Target As Range)

userForm_naa.TextBox_n_efector.Text = Cells(Target.Row, Target.Column - 11).Value
userForm_naa.TextBox_denominacion_efector.Text = Cells(Target.Row, Target.Column - 10).Value
userForm_naa.TextBox_orden_pago_factura = Cells(Target.Row, Target.Column - 9).Value
userForm_naa.TextBox_beneficiario.Text = Cells(Target.Row, Target.Column - 8).Value & " " & Cells(Target.Row, Target.Column - 7).Value
userForm_naa.TextBox_documento.Text = Cells(Target.Row, Target.Column - 6).Value
userForm_naa.TextBox_clave_beneficiario.Text = Cells(Target.Row, Target.Column - 5).Value
userForm_naa.TextBox_codigo.Text = Cells(Target.Row, Target.Column - 3).Value
userForm_naa.TextBox_descripcion.Text = Cells(Target.Row, Target.Column - 2).Value
userForm_naa.TextBox_fecha_prestacion.Text = Cells(Target.Row, Target.Column - 1).Value
userForm_naa.TextBox_fecha_nacimiento.Text = Cells(Target.Row, Target.Column + 3).Value
userForm_naa.TextBox_monto = Cells(Target.Row, Target.Column + 19).Value
userForm_naa.TextBox_edad.Text = Cells(Target.Row, Target.Column + 21).Value


'verifica que si la celda de nombre de prestacion contiene "La prestacion no corresponde al grupo poblacional" si
'esto es verdadero le otorga un valor al textbox de control de fuente de informacion
If (Cells(filaDobleClick, columnaDobleClick - 2).Value = "La prestación no corresponde al grupo poblacional") Then

    userForm_naa.dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"
    
End If

End Sub


'brief: copia los datos ya relevados a un userform
'param: rango donde se hizo doble click
'return: void

Sub userForm_naa_copiar_datos_relevamiento(ByVal Target As Range)


Dim leyenda As String

leyenda = "Dato no obligatorio"


userForm_naa.dato_fuente.Text = Cells(Target.Row, Target.Column + 1).Value


'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
If (userForm_naa.dato_fuente.Text <> "No consta fuente de información" And userForm_naa.dato_fuente.Text <> "Prestación inexistente" And _
userForm_naa.dato_control_fuente.Text <> "Fuente invalida") Then

    userForm_naa.dato_diagnostico.Text = Cells(Target.Row, Target.Column + 9).Value
    
    'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
    'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
    'lineas de codigo de mas
    userForm_naa.dato_estudios.Text = Cells(Target.Row, Target.Column + 13).Value
    If (userForm_naa.dato_estudios.Text = "") Then
    userForm_naa.dato_estudios.Text = leyenda
    End If
    
    userForm_naa.dato_evaluacion_riesgo.Text = Cells(Target.Row, Target.Column + 14).Value
    If (userForm_naa.dato_evaluacion_riesgo.Text = "") Then
    userForm_naa.dato_evaluacion_riesgo.Text = leyenda
    End If
    
    userForm_naa.dato_ta.Text = Cells(Target.Row, Target.Column + 7).Value
    If (userForm_naa.dato_ta.Text = "") Then
    userForm_naa.dato_ta.Text = leyenda
    End If
    
    userForm_naa.dato_imc.Text = Cells(Target.Row, Target.Column + 8).Value
    If (userForm_naa.dato_imc.Text = "") Then
    userForm_naa.dato_imc.Text = leyenda
    End If
    
    userForm_naa.dato_percentilo.Text = Cells(Target.Row, Target.Column + 6).Value
    If (userForm_naa.dato_percentilo.Text = "") Then
    userForm_naa.dato_percentilo.Text = leyenda
    End If
    
    userForm_naa.dato_peso.Text = Cells(Target.Row, Target.Column + 4).Value
    If (userForm_naa.dato_peso.Text = "") Then
    userForm_naa.dato_peso.Text = leyenda
    End If
    
    userForm_naa.dato_talla.Text = Cells(Target.Row, Target.Column + 5).Value
    If (userForm_naa.dato_talla.Text = "") Then
    userForm_naa.dato_talla.Text = leyenda
    End If
    
    userForm_naa.dato_plan_seguimiento.Text = Cells(Target.Row, Target.Column + 12).Value
    If (userForm_naa.dato_plan_seguimiento.Text = "") Then
    userForm_naa.dato_plan_seguimiento.Text = leyenda
    End If
    
    userForm_naa.dato_tratamiento_instaurado.Text = Cells(Target.Row, Target.Column + 10).Value
    If (userForm_naa.dato_tratamiento_instaurado.Text = "") Then
    userForm_naa.dato_tratamiento_instaurado.Text = leyenda
    End If
    
    userForm_naa.dato_transcripcion.Text = Cells(Target.Row, Target.Column + 11).Value
    If (userForm_naa.dato_transcripcion.Text = "") Then
    userForm_naa.dato_transcripcion.Text = leyenda
    End If
    
    userForm_naa.dato_constancia_inmunizaciones.Text = Cells(Target.Row, Target.Column + 18).Value
    If (userForm_naa.dato_constancia_inmunizaciones.Text = "") Then
    userForm_naa.dato_constancia_inmunizaciones.Text = leyenda
    End If
    
    userForm_naa.dato_fecha_notificacion.Text = Cells(Target.Row, Target.Column + 15).Value
    If (userForm_naa.dato_fecha_notificacion.Text = "") Then
    userForm_naa.dato_fecha_notificacion.Text = leyenda
    End If
    
    userForm_naa.dato_firma.Text = Cells(Target.Row, Target.Column + 16).Value
    If (userForm_naa.dato_firma.Text = "") Then
    userForm_naa.dato_firma.Text = leyenda
    End If
    
    userForm_naa.dato_sello.Text = Cells(Target.Row, Target.Column + 17).Value
    If (userForm_naa.dato_sello.Text = "") Then
    userForm_naa.dato_sello.Text = leyenda
    End If
    
    'userForm_naa.dato_control_fuente.Text = Cells(target.Row, target.Column + 2).Value
    
    'userForm_naa.dato_validacion.Text = Cells(target.Row, target.Column + 21).Value

Else
    
    Call userForm_naa_dato_no_obligatorio
    
End If

userForm_naa.dato_observaciones.Text = Cells(Target.Row, Target.Column + 20).Value

End Sub


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_naa_bloquear()

userForm_naa.TextBox_n_efector.Locked = True
userForm_naa.TextBox_denominacion_efector.Locked = True
userForm_naa.TextBox_orden_pago_factura.Locked = True
userForm_naa.TextBox_beneficiario.Locked = True
userForm_naa.TextBox_clave_beneficiario.Locked = True
userForm_naa.TextBox_documento.Locked = True
userForm_naa.TextBox_codigo.Locked = True
userForm_naa.TextBox_fecha_prestacion.Locked = True
userForm_naa.TextBox_fecha_nacimiento.Locked = True
userForm_naa.TextBox_edad.Locked = True
userForm_naa.TextBox_monto.Locked = True
userForm_naa.TextBox_descripcion.Locked = True

'userform_naa.dato_fuente.Locked = True
'userform_naa.dato_diagnostico.Locked = True
userForm_naa.dato_estudios.Locked = True
userForm_naa.dato_evaluacion_riesgo.Locked = True
userForm_naa.dato_ta.Locked = True
userForm_naa.dato_imc.Locked = True
userForm_naa.dato_percentilo.Locked = True
userForm_naa.dato_peso.Locked = True
userForm_naa.dato_talla.Locked = True
userForm_naa.dato_plan_seguimiento.Locked = True
userForm_naa.dato_tratamiento_instaurado.Locked = True
userForm_naa.dato_transcripcion.Locked = True
userForm_naa.dato_constancia_inmunizaciones.Locked = True
userForm_naa.dato_fecha_notificacion.Locked = True
userForm_naa.dato_firma.Locked = True
userForm_naa.dato_sello.Locked = True
'userform_naa.dato31.Locked = True
'userform_naa.dato32.Locked = True
'userform_naa.dato_validacion.Locked = True


End Function


'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userForm_naa_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

userForm_naa.dato_diagnostico.Text = leyenda
userForm_naa.dato_estudios.Text = leyenda
userForm_naa.dato_evaluacion_riesgo.Text = leyenda
userForm_naa.dato_ta.Text = leyenda
userForm_naa.dato_imc.Text = leyenda
userForm_naa.dato_percentilo.Text = leyenda
userForm_naa.dato_peso.Text = leyenda
userForm_naa.dato_talla.Text = leyenda
userForm_naa.dato_plan_seguimiento.Text = leyenda
userForm_naa.dato_tratamiento_instaurado.Text = leyenda
userForm_naa.dato_transcripcion.Text = leyenda
userForm_naa.dato_constancia_inmunizaciones.Text = leyenda
userForm_naa.dato_fecha_notificacion.Text = leyenda
userForm_naa.dato_firma.Text = leyenda
userForm_naa.dato_sello.Text = leyenda

userForm_naa.dato_diagnostico.BackColor = RGB(169, 169, 169)
userForm_naa.dato_estudios.BackColor = RGB(169, 169, 169)
userForm_naa.dato_evaluacion_riesgo.BackColor = RGB(169, 169, 169)
userForm_naa.dato_ta.BackColor = RGB(169, 169, 169)
userForm_naa.dato_imc.BackColor = RGB(169, 169, 169)
userForm_naa.dato_percentilo.BackColor = RGB(169, 169, 169)
userForm_naa.dato_peso.BackColor = RGB(169, 169, 169)
userForm_naa.dato_talla.BackColor = RGB(169, 169, 169)
userForm_naa.dato_plan_seguimiento.BackColor = RGB(169, 169, 169)
userForm_naa.dato_tratamiento_instaurado.BackColor = RGB(169, 169, 169)
userForm_naa.dato_transcripcion.BackColor = RGB(169, 169, 169)
userForm_naa.dato_constancia_inmunizaciones.BackColor = RGB(169, 169, 169)
userForm_naa.dato_fecha_notificacion.BackColor = RGB(169, 169, 169)
userForm_naa.dato_firma.BackColor = RGB(169, 169, 169)
userForm_naa.dato_sello.BackColor = RGB(169, 169, 169)

userForm_naa.dato_diagnostico.Locked = True
userForm_naa.dato_estudios.Locked = True
userForm_naa.dato_evaluacion_riesgo.Locked = True
userForm_naa.dato_ta.Locked = True
userForm_naa.dato_imc.Locked = True
userForm_naa.dato_percentilo.Locked = True
userForm_naa.dato_peso.Locked = True
userForm_naa.dato_talla.Locked = True
userForm_naa.dato_plan_seguimiento.Locked = True
userForm_naa.dato_tratamiento_instaurado.Locked = True
userForm_naa.dato_transcripcion.Locked = True
userForm_naa.dato_constancia_inmunizaciones.Locked = True
userForm_naa.dato_fecha_notificacion.Locked = True
userForm_naa.dato_firma.Locked = True
userForm_naa.dato_sello.Locked = True

End Function


'brief: copia los datos del userform_naa al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_naa_guardar_datos(ByVal fila As Integer)

With userForm_naa

    Cells(fila, 14).Value = .dato_fuente.Text
    Cells(fila, 22).Value = .dato_diagnostico.Text
    Cells(fila, 26).Value = .dato_estudios.Text
    Cells(fila, 27).Value = .dato_evaluacion_riesgo.Text
    Cells(fila, 20).Value = .dato_ta.Text
    Cells(fila, 21).Value = .dato_imc.Text
    Cells(fila, 19).Value = .dato_percentilo.Text
    Cells(fila, 17).Value = .dato_peso.Text
    Cells(fila, 18).Value = .dato_talla.Text
    Cells(fila, 25).Value = .dato_plan_seguimiento.Text
    Cells(fila, 23).Value = .dato_tratamiento_instaurado.Text
    Cells(fila, 24).Value = .dato_transcripcion.Text
    Cells(fila, 31).Value = .dato_constancia_inmunizaciones.Text
    Cells(fila, 28).Value = .dato_fecha_notificacion.Text
    Cells(fila, 29).Value = .dato_firma.Text
    Cells(fila, 30).Value = .dato_sello.Text
    Cells(fila, 33).Value = .dato_observaciones.Text

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


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_naa_verificacion_blancos() As Integer
 
If (userForm_naa.dato_fuente.Text = "" Or userForm_naa.dato_diagnostico.Text = "" Or userForm_naa.dato_estudios.Text = "" Or userForm_naa.dato_evaluacion_riesgo.Text = "" Or _
userForm_naa.dato_ta.Text = "" Or userForm_naa.dato_imc.Text = "" Or userForm_naa.dato_percentilo.Text = "" Or userForm_naa.dato_peso.Text = "" _
Or userForm_naa.dato_talla.Text = "" Or userForm_naa.dato_plan_seguimiento.Text = "" Or userForm_naa.dato_tratamiento_instaurado.Text = "" Or userForm_naa.dato_transcripcion.Text = "" _
Or userForm_naa.dato_constancia_inmunizaciones.Text = "" Or userForm_naa.dato_fecha_notificacion.Text = "" Or userForm_naa.dato_firma.Text = "" Or userForm_naa.dato_sello.Text = "") Then

userForm_naa_verificacion_blancos = 1

Else

userForm_naa_verificacion_blancos = 0

End If


End Function

'brief desbloquea y limpia los combobox y textbos que corresponde a los datos obligatorios
'param void
'return void

Sub userForm_naa_permitir_campos_requeridos(ByVal codigo As String)

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
If (userForm_naa.dato_validacion.Text <> leyenda And userForm_naa.dato_validacion.Text <> leyenda2 And _
userForm_naa.dato_validacion.Text <> leyenda3) Then

    If (userForm_naa.dato_diagnostico.Text = leyenda4) Then
    
    userForm_naa.dato_diagnostico.Locked = False
    userForm_naa.dato_diagnostico.BackColor = RGB(255, 255, 255)
    userForm_naa.dato_diagnostico.Text = ""
    
    End If
    
End If


'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_naa.dato_validacion.Text <> leyenda And userForm_naa.dato_validacion.Text <> leyenda2 And userForm_naa.dato_validacion.Text <> leyenda3) Then


    For i = 1 To 250
    
        'coincidencia de poblacion con el codigo o el codigo con la poblacion embarazo
        If ((ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 32).Value And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo) Or (ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = "Embarazos")) Then
                
                've si el codigo es CTC001A97 y la poblacion es niños (distintos datos obligatorios por edad)
                If (codigo = "CTC001A97" And ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 32).Value = "Niños") Then
                    
                    'hace coincidir la edad en el relevamiento con la descripcion
                    If ((ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 21).Value < 1 And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "menores de 1") Or _
                    ((ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 21).Value >= 1 And ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 21).Value < 6) And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "1 a 5") Or _
                    (ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 21).Value >= 6 And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "mayores a 6")) Then

                        For j = 5 To 31
                            
                            'verifica que que el dato que corresponde a la celda se obligatorio
                            If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                                
                                'copia el nombre del dato obligatorio y empiesa a verificar cual es
                                texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                                
                                Select Case texto
                                
                                    Case "Peso"
                                    
                                        With userForm_naa.dato_peso
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Talla"
                                    
                                        With userForm_naa.dato_talla
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "TA"
                                        
                                        With userForm_naa.dato_ta
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "IMC"
                                    
                                        With userForm_naa.dato_imc
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Informe o transcripción de estudios solicitados"
                                        
                                        With userForm_naa.dato_transcripcion
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                    
                                    
                                    Case "Evaluacion de riesgo"
                                    
                                        With userForm_naa.dato_evaluacion_riesgo
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Diagnostico"
                                        
                                        With userForm_naa.dato_diagnostico
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Tratamiento instaurado"
                                        
                                        With userForm_naa.dato_tratamiento_instaurado
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Plan de seguimiento"
                                    
                                        With userForm_naa.dato_plan_seguimiento
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                    
                                    
                                    Case "Constancia de aplicación de inmunizaciones"
                                    
                                        With userForm_naa.dato_constancia_inmunizaciones
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    Case "Fecha de notificación o de parto/cesárea"
                                    
                                        With userForm_naa.dato_fecha_notificacion
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    Case "Firma"
                                        
                                        With userForm_naa.dato_firma
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    
                                    Case "Sello"
                                        
                                        With userForm_naa.dato_sello
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
'                                    Case "Altura uterina"
'
'                                        With userForm_naa
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
'
'
'                                    Case "Apgar"
'
'                                        With userForm_naa
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
'
'
'                                    Case "Calculo de amenorrea"
'
'                                        With userForm_naa.dato_amen
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
'
'
'                                    Case "Constancia de solicitud de screening neonatal"
'
'                                        With userForm_naa.dato_peso
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
'
'
'                                    Case "Diagnostico de embarazo de alto riesgo"
'
'                                        With userForm_naa.dato_peso
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End Wit
'
'
'                                    Case "Perimetro cefalico"
'
'                                        With userForm_naa.dato_perimetro_cefalico
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
                                        
                                        
                                    Case "Percentilo"
                                    
                                        With userForm_naa.dato_percentilo
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                    
'                                    Case "N de Control Prenatal"
'
'                                        With userForm_naa.dato_control_prenatal
'                                        .Locked = False
'                                        If (.Text = leyenda4) Then
'                                            .Text = ""
'                                         End If
'                                        .BackColor = RGB(255, 255, 255)
'                                        End With
                                        
                                    Case "Examen mamario", "Evaluacion genitourinaria", "Odontograma", "Medicion de agudeza visual"
                                    
                                        With userForm_naa.dato_estudios
                                            .Locked = False
                                            If (.Text = "Dato no obligatorio") Then
                                                .Text = ""
                                            End If
                                            .BackColor = RGB(255, 255, 255)
                                        End With

                                End Select
                                
                            End If

                        Next
                        
                        'para romper el primer for (el de i)
                        i = 250
                    
                    End If
                    
                
                
                Else
                
                    For j = 5 To 31
                        
                        'verifica que que el dato que corresponde a la celda se obligatorio
                        If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                            
                            'copia el nombre del dato obligatorio y empiesa a verificar cual es
                            texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                            
                            Select Case texto
                                
                                Case "Peso"
                                
                                    With userForm_naa.dato_peso
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "Talla"
                                
                                    With userForm_naa.dato_talla
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "TA"
                                    
                                    With userForm_naa.dato_ta
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "IMC"
                                
                                    With userForm_naa.dato_imc
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "Informe o transcripción de estudios solicitados"
                                    
                                    With userForm_naa.dato_transcripcion
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                
                                
                                Case "Evaluacion de riesgo"
                                
                                    With userForm_naa.dato_evaluacion_riesgo
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "Diagnostico"
                                    
                                    With userForm_naa.dato_diagnostico
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "Tratamiento instaurado"
                                    
                                    With userForm_naa.dato_tratamiento_instaurado
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "Plan de seguimiento"
                                
                                    With userForm_naa.dato_plan_seguimiento
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                
                                
                                Case "Constancia de aplicación de inmunizaciones"
                                
                                    With userForm_naa.dato_constancia_inmunizaciones
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                Case "Fecha de notificación o de parto/cesárea"
                                    
                                    With userForm_naa.dato_fecha_notificacion
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
                                Case "Firma"
                                    
                                    With userForm_naa.dato_firma
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                
                                Case "Sello"
                                    
                                    With userForm_naa.dato_sello
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                    
'                                Case "Altura uterina"
'
'                                    With userForm_naa
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                        .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End With
'
'
'                                Case "Apgar"
'
'                                    With userForm_naa
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                        .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End With
'
'
'                                Case "Calculo de amenorrea"
'
'                                    With userForm_naa.dato_amen
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                        .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End With
'
'
'                                Case "Constancia de solicitud de screening neonatal"
'
'                                    With userForm_naa.dato_peso
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                        .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End With
'
'
'                                Case "Diagnostico de embarazo de alto riesgo"
'
'                                    With userForm_naa.dato_peso
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                        .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End Wit
'
'
'                                Case "Perimetro cefalico"
'
'                                    With userForm_naa.dato_perimetro_cefalico
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                        .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End With
                                    
                                    
                                Case "Percentilo"
                                
                                    With userForm_naa.dato_percentilo
                                    .Locked = False
                                    If (.Text = leyenda4) Then
                                        .Text = ""
                                    End If
                                    .BackColor = RGB(255, 255, 255)
                                    End With
                                    
                                
'                                Case "N de Control Prenatal"
'
'                                    With userForm_naa.dato_control_prenatal
'                                    .Locked = False
'                                    If (.Text = leyenda4) Then
'                                         .Text = ""
'                                     End If
'                                    .BackColor = RGB(255, 255, 255)
'                                    End With

                                Case "Examen mamario", "Evaluacion genitourinaria", "Odontograma", "Medicion de agudeza visual"
                                    
                                    With userForm_naa.dato_estudios
                                        .Locked = False
                                        If (.Text = "Dato no obligatorio") Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                    End With
                                    
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










