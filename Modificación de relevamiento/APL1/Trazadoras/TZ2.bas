Attribute VB_Name = "TZ2"
'declaracion de variables globales

Public filaDobleClickTz2 As Integer

Public columnaDobleClickTz2 As Integer

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteTz2 As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorTz2 As Integer


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_tz2_bloquear()

With userForm_tz2

    .TextBox_n_efector.Locked = True
    .TextBox_denominacion_efector.Locked = True
    .TextBox_beneficiario.Locked = True
    .TextBox_documento.Locked = True
    .TextBox_fecha_nacimiento.Locked = True
    .dato_fecha_control_1.Locked = True
    .dato_fecha_control_2.Locked = True
    .dato_fecha_control_3.Locked = True
    .dato_fecha_control_4.Locked = True
    
    .dato_fum.Locked = True
    .dato_control_1_completo.Locked = True
    .dato_control_2_completo.Locked = True
    .dato_control_3_completo.Locked = True
    .dato_control_4_completo.Locked = True
    .dato_edad_gestacional.Locked = True
    

End With

End Function


'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_tz2_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz2
    
    .dato_fum.Text = leyenda
    .dato_control_1_completo.Text = leyenda
    .dato_control_2_completo.Text = leyenda
    .dato_control_3_completo.Text = leyenda
    .dato_control_4_completo.Text = leyenda
    .dato_edad_gestacional.Text = leyenda
    
    .dato_fum.BackColor = RGB(169, 169, 169)
    .dato_control_1_completo.BackColor = RGB(169, 169, 169)
    .dato_control_2_completo.BackColor = RGB(169, 169, 169)
    .dato_control_3_completo.BackColor = RGB(169, 169, 169)
    .dato_control_4_completo.BackColor = RGB(169, 169, 169)
    .dato_edad_gestacional.BackColor = RGB(169, 169, 169)
    
    .dato_fum.Locked = True
    .dato_control_1_completo.Locked = True
    .dato_control_2_completo.Locked = True
    .dato_control_3_completo.Locked = True
    .dato_control_4_completo.Locked = True
    .dato_edad_gestacional.Locked = True
    
End With

End Function


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_tz2_verificacion_blancos() As Integer

With userForm_tz2

    If (.dato_fuente.Text = "" Or .dato_fum.Text = "" Or .dato_fecha_control_1.Text = "" Or .dato_control_1_completo.Text = "" Or _
    .dato_fecha_control_2.Text = "" Or .dato_control_2_completo.Text = "" Or .dato_fecha_control_3.Text = "" Or _
    .dato_control_3_completo = "" Or .dato_fecha_control_4.Text = "" Or .dato_control_4_completo.Text = "" Or .dato_edad_gestacional.Text = "") Then
        
        userForm_tz2_verificacion_blancos = 1

    Else

        userForm_tz2_verificacion_blancos = 0
    
    End If
 
End With

End Function




'brief: copia los datos del userform_tz2 al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_tz2_guardar_datos(ByVal fila As Integer)

With userForm_tz2
    
    Cells(fila, 10).Value = .dato_fuente.Text
    Cells(fila, 12).Value = .dato_fum.Text
    Cells(fila, 14).Value = .dato_control_1_completo.Text
    Cells(fila, 16).Value = .dato_control_2_completo.Text
    Cells(fila, 18).Value = .dato_control_3_completo.Text
    Cells(fila, 20).Value = .dato_control_4_completo.Text
    Cells(fila, 21).Value = .dato_edad_gestacional.Text
    Cells(fila, 22).Value = .dato_observaciones.Text
    
    'para que el auditor pueda filtrar por A o B para completar el acta
    If (.dato_fuente.Text = "No consta fuente de información") Then
        Cells(fila, 11).Value = "A"
        
    ElseIf (.dato_fuente.Text = "Prestación inexistente") Then
        Cells(fila, 11).Value = "B"
        
    Else
        Cells(fila, 11).Value = "No labrar acta"
        
    End If
    
End With

End Sub


'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub copiar_tz2_datos_fijos(ByVal fila As Integer)

With userForm_tz2


    .TextBox_n_efector.Text = Cells(fila, 3).Value
    .TextBox_denominacion_efector.Text = Cells(fila, 4).Value
    .TextBox_documento.Text = Cells(fila, 5).Value
    .TextBox_beneficiario.Text = Cells(fila, 6).Value & " " & Cells(fila, 7).Value
    .TextBox_fecha_nacimiento.Text = Cells(fila, 8).Value
    .dato_fecha_control_1.Text = Cells(fila, 13).Value
    .dato_fecha_control_2.Text = Cells(fila, 15).Value
    .dato_fecha_control_3.Text = Cells(fila, 17).Value
    .dato_fecha_control_4.Text = Cells(fila, 19).Value
    
End With

End Sub


'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_tz2_copiar_datos_relevamiento(ByVal fila As Integer)

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz2
    
    .dato_fuente.Text = Cells(fila, 10).Value
    
    'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
    If (.dato_fuente.Text <> "No consta fuente de información" And .dato_fuente.Text <> "Prestación inexistente") Then
    
        'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
        'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
        'lineas de codigo de mas
    
        .dato_fum.Text = Cells(fila, 12).Value
        If (.dato_fum.Text = "") Then
            .dato_fum.Text = leyenda
        End If
        
        .dato_control_1_completo.Text = Cells(fila, 14).Value
        If (.dato_control_1_completo.Text = "") Then
            .dato_control_1_completo.Text = leyenda
        End If
        
        .dato_control_2_completo.Text = Cells(fila, 16).Value
        If (.dato_control_2_completo.Text = "") Then
            .dato_control_2_completo.Text = leyenda
        End If
        
        .dato_control_3_completo.Text = Cells(fila, 18).Value
        If (.dato_control_3_completo.Text = "") Then
            .dato_control_3_completo.Text = leyenda
        End If
        
        .dato_control_4_completo.Text = Cells(fila, 20).Value
        If (.dato_control_4_completo.Text = "") Then
            .dato_control_4_completo.Text = leyenda
        End If
        
        .dato_edad_gestacional.Text = Cells(fila, 21).Value
        If (.dato_edad_gestacional.Text = "") Then
            .dato_edad_gestacional.Text = leyenda
        End If
        
    Else
    
        Call userform_tz2_dato_no_obligatorio
    
    End If
    
    .dato_observaciones.Text = Cells(fila, 22).Value
    
End With

End Sub



'brief desbloquea y limpia los comboboxs y textboxs que corresponde a los datos obligatorios
'param void
'return void

Function userForm_tz2_permitir_campos_requeridos()

Dim leyenda As String
Dim leyenda2 As String
Dim leyenda3 As String

leyenda = "Labrar acta"
leyenda2 = "Labrar acta e indicar fuente de información en observaciones"
leyenda3 = "Dato no obligatorio"

'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_tz2.dato_validacion.Text <> leyenda And userForm_tz2.dato_validacion.Text <> leyenda2) Then
    
    With userForm_tz2
                    
        With .dato_fum
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_control_1_completo
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_control_2_completo
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_control_3_completo
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_control_4_completo
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_edad_gestacional
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
    End With
    
End If

End Function
