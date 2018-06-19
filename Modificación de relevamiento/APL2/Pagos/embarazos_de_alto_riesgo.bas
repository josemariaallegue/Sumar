Attribute VB_Name = "embarazos_de_alto_riesgo"
'declaracion de variables globales

Public filaDobleClickEAR As Integer

Public columnaDobleClickEAR As Integer

Public codigoDobleClickEAR As String

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteEAR As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorEAR As Integer

'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub copiar_ear_datos_fijos(ByVal Target As Range)

With userForm_ear

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
    .TextBox_monto = Cells(Target.Row, Target.Column + 11).Value
    .TextBox_edad.Text = Cells(Target.Row, Target.Column + 13).Value

End With

'verifica que si la celda de nombre de prestacion contiene "La prestacion no corresponde al grupo poblacional" si
'esto es verdadero le otorga un valor al textbox de control de fuente de informacion
If (Cells(filaDobleClickEAR, columnaDobleClickEAR - 2).Value = "La prestación no corresponde al grupo poblacional") Then

    userForm_ear.dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"
    
End If

End Sub


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_ear_bloquear()

With userForm_ear

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
    
    .dato_motivo_egreso.Locked = True
    .dato_reporte.Locked = True
    .dato_fecha_ingreso.Locked = True
    .dato_tratamiento_instaurado.Locked = True
    .dato_fecha_egreso.Locked = True
    .dato_uti.Locked = True
    .dato_sala.Locked = True
    .dato_fecha_notificacion.Locked = True
    .dato_firma.Locked = True
    .dato_sello.Locked = True

End With

End Function


'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_ear_copiar_datos_relevamiento(ByVal Target As Range)

Dim leyenda As String

leyenda = "Dato no obligatorio"


userForm_ear.dato_fuente.Text = Cells(Target.Row, Target.Column + 1).Value

'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
If (userForm_ear.dato_fuente.Text <> "No consta fuente de información" And userForm_ear.dato_fuente.Text <> "Prestación inexistente" And _
userForm_ear.dato_control_fuente.Text <> "Fuente invalida") Then

    userForm_ear.dato_diagnostico.Text = Cells(Target.Row, Target.Column + 6).Value
    
    'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
    'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
    'lineas de codigo de mas
    userForm_ear.dato_reporte.Text = Cells(Target.Row, Target.Column + 4).Value
    If (userForm_ear.dato_reporte.Text = "") Then
        userForm_ear.dato_reporte.Text = leyenda
    End If
    
    userForm_ear.dato_fecha_ingreso.Text = Cells(Target.Row, Target.Column + 5).Value
    If (userForm_ear.dato_fecha_ingreso.Text = "") Then
        userForm_ear.dato_fecha_ingreso.Text = leyenda
    End If
    
    userForm_ear.dato_tratamiento_instaurado.Text = Cells(Target.Row, Target.Column + 7).Value
    If (userForm_ear.dato_tratamiento_instaurado.Text = "") Then
        userForm_ear.dato_tratamiento_instaurado.Text = leyenda
    End If
    
    userForm_ear.dato_fecha_egreso.Text = Cells(Target.Row, Target.Column + 8).Value
    If (userForm_ear.dato_fecha_egreso.Text = "") Then
        userForm_ear.dato_fecha_egreso.Text = leyenda
    End If
    
    userForm_ear.dato_uti.Text = Cells(Target.Row, Target.Column + 9).Value
    If (userForm_ear.dato_uti.Text = "") Then
        userForm_ear.dato_uti.Text = leyenda
    End If
    
    userForm_ear.dato_sala.Text = Cells(Target.Row, Target.Column + 10).Value
    If (userForm_ear.dato_sala.Text = "") Then
        userForm_ear.dato_sala.Text = leyenda
    End If
    
    userForm_ear.dato_motivo_egreso.Text = Cells(Target.Row, Target.Column + 11).Value
    If (userForm_ear.dato_motivo_egreso.Text = "") Then
        userForm_ear.dato_motivo_egreso.Text = leyenda
    End If
    
    userForm_ear.dato_fecha_notificacion.Text = Cells(Target.Row, Target.Column + 12).Value
    If (userForm_ear.dato_fecha_notificacion.Text = "") Then
        userForm_ear.dato_fecha_notificacion.Text = leyenda
    End If
    
    userForm_ear.dato_firma.Text = Cells(Target.Row, Target.Column + 13).Value
    If (userForm_ear.dato_firma.Text = "") Then
        userForm_ear.dato_firma.Text = leyenda
    End If
    
    userForm_ear.dato_sello.Text = Cells(Target.Row, Target.Column + 14).Value
    If (userForm_ear.dato_sello.Text = "") Then
        userForm_ear.dato_sello.Text = leyenda
    End If
    
    'userForm_ear.dato_control_fuente.Text = Cells(target.Row, target.Column + 2).Value
    
    'userForm_ear.dato_validacion.Text = Cells(target.Row, target.Column + 21).Value
    
Else

    Call userform_ear_dato_no_obligatorio

End If

userForm_ear.dato_observaciones.Text = Cells(Target.Row, Target.Column + 16).Value

End Sub

'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_ear_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_ear
    
    .dato_diagnostico.Text = leyenda
    .dato_motivo_egreso.Text = leyenda
    .dato_reporte.Text = leyenda
    .dato_fecha_ingreso.Text = leyenda
    .dato_fecha_egreso.Text = leyenda
    .dato_tratamiento_instaurado.Text = leyenda
    .dato_fecha_egreso.Text = leyenda
    .dato_uti.Text = leyenda
    .dato_sala.Text = leyenda
    .dato_fecha_notificacion = leyenda
    .dato_firma.Text = leyenda
    .dato_sello.Text = leyenda
    
    .dato_diagnostico.BackColor = RGB(169, 169, 169)
    .dato_motivo_egreso.BackColor = RGB(169, 169, 169)
    .dato_reporte.BackColor = RGB(169, 169, 169)
    .dato_fecha_ingreso.BackColor = RGB(169, 169, 169)
    .dato_fecha_egreso.BackColor = RGB(169, 169, 169)
    .dato_tratamiento_instaurado.BackColor = RGB(169, 169, 169)
    .dato_fecha_egreso.BackColor = RGB(169, 169, 169)
    .dato_uti.BackColor = RGB(169, 169, 169)
    .dato_sala.BackColor = RGB(169, 169, 169)
    .dato_fecha_notificacion.BackColor = RGB(169, 169, 169)
    .dato_firma.BackColor = RGB(169, 169, 169)
    .dato_sello.BackColor = RGB(169, 169, 169)
  
    .dato_diagnostico.Locked = True
    .dato_motivo_egreso.Locked = True
    .dato_reporte.Locked = True
    .dato_fecha_ingreso.Locked = True
    .dato_fecha_egreso.Locked = True
    .dato_tratamiento_instaurado.Locked = True
    .dato_fecha_egreso.Locked = True
    .dato_uti.Locked = True
    .dato_sala.Locked = True
    .dato_fecha_notificacion = True
    .dato_firma.Locked = True
    .dato_sello.Locked = True
    
End With

End Function


'brief: copia los datos del userform_ear al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_ear_guardar_datos(ByVal fila As Integer)

With userForm_ear
    
    Cells(fila, 14).Value = .dato_fuente.Text
    Cells(fila, 17).Value = .dato_reporte.Text
    Cells(fila, 18).Value = .dato_fecha_ingreso.Text
    Cells(fila, 19).Value = .dato_diagnostico.Text
    Cells(fila, 20).Value = .dato_tratamiento_instaurado.Text
    Cells(fila, 21).Value = .dato_fecha_egreso.Text
    Cells(fila, 22).Value = .dato_uti.Text
    Cells(fila, 23).Value = .dato_sala.Text
    Cells(fila, 24).Value = .dato_motivo_egreso.Text
    Cells(fila, 25).Value = .dato_fecha_notificacion.Text
    Cells(fila, 26).Value = .dato_firma.Text
    Cells(fila, 27).Value = .dato_sello.Text
    Cells(fila, 29).Value = .dato_observaciones.Text
    
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

Sub userForm_ear_permitir_campos_requeridos(ByVal codigo As String)

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
If (userForm_ear.dato_validacion.Text <> leyenda And userForm_ear.dato_validacion.Text <> leyenda2 And _
userForm_ear.dato_validacion.Text <> leyenda3) Then

    If (userForm_ear.dato_diagnostico.Text = leyenda4) Then
    
        userForm_ear.dato_diagnostico.Locked = False
        userForm_ear.dato_diagnostico.BackColor = RGB(255, 255, 255)
        userForm_ear.dato_diagnostico.Text = ""
    
    End If
    
End If

'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_ear.dato_validacion.Text <> leyenda And userForm_ear.dato_validacion.Text <> leyenda2 And userForm_ear.dato_validacion.Text <> leyenda3) Then


    For i = 1 To 250
    
        'coincidencia de poblacion con el codigo o el codigo con la poblacion embarazo
        If ((ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = ActiveSheet.Cells(filaDobleClickEAR, columnaDobleClickEAR + 27).Value And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo) Or (ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = "Embarazos")) Then
                
            With userForm_ear
            
                With .dato_motivo_egreso
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_reporte
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_fecha_ingreso
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_tratamiento_instaurado
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_fecha_egreso
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_uti
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_sala
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_fecha_notificacion
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_firma
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
                With .dato_sello
                    .Locked = False
                    If (.Text = leyenda4) Then
                        .Text = ""
                    End If
                    .BackColor = RGB(255, 255, 255)
                End With
                
            End With
            
            'para terminar el for
            i = 250
        
        End If
        
    Next
       
End If

End Sub


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_ear_verificacion_blancos() As Integer

With userForm_ear

    If (.dato_fuente.Text = "" Or .dato_diagnostico.Text = "" Or .dato_motivo_egreso.Text = "" Or .dato_reporte.Text = "" Or _
    .dato_fecha_ingreso.Text = "" Or .dato_tratamiento_instaurado.Text = "" Or .dato_fecha_egreso.Text = "" Or .dato_uti.Text = "" Or _
    .dato_sala.Text = "" Or .dato_fecha_notificacion.Text = "" Or .dato_firma.Text = "" Or .dato_sello.Text = "") Then
        
        userForm_ear_verificacion_blancos = 1

    Else

        userForm_ear_verificacion_blancos = 0
    
    End If
 
End With

End Function


