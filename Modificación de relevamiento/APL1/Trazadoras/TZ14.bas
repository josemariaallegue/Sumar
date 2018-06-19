Attribute VB_Name = "TZ14"
'declaracion de variables globales

Public filaDobleClickTz14 As Integer

Public columnaDobleClickTz14 As Integer

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteTz14 As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorTz14 As Integer


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_tz14_bloquear()

With userForm_tz14

    .TextBox_beneficiario.Locked = True
    .TextBox_documento.Locked = True
    .TextBox_fecha_obito.Locked = True
    .TextBox_fecha_comite.Locked = True
    
    .dato_fecha_comite_pregunta.Locked = True
    .dato_fecha_comite_terreno.Locked = True
    .dato_diagnostico.Locked = True
    

End With

End Function


'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_tz14_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz14
    
    .dato_fecha_comite_pregunta.Text = leyenda
    .dato_fecha_comite_terreno.Text = leyenda
    .dato_diagnostico.Text = leyenda
    
    .dato_fecha_comite_pregunta.BackColor = RGB(169, 169, 169)
    .dato_fecha_comite_terreno.BackColor = RGB(169, 169, 169)
    .dato_diagnostico.BackColor = RGB(169, 169, 169)
    
    .dato_fecha_comite_pregunta.Locked = True
    .dato_fecha_comite_terreno.Locked = True
    .dato_diagnostico.Locked = True
    
End With

End Function


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_tz14_verificacion_blancos() As Integer

With userForm_tz14

    If (.dato_fecha_comite_pregunta.Text = "" Or .dato_fecha_comite_terreno.Text = "" Or .dato_diagnostico.Text = "") Then
        
        userForm_tz14_verificacion_blancos = 1

    Else

        userForm_tz14_verificacion_blancos = 0
    
    End If
 
End With

End Function




'brief: copia los datos del userform_tz14 al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_tz14_guardar_datos(ByVal fila As Integer)

With userForm_tz14
    
    Cells(fila, 8).Value = .dato_fecha_comite_pregunta.Text
    Cells(fila, 9).Value = .dato_fecha_comite_terreno.Text
    Cells(fila, 10).Value = .dato_diagnostico.Text
    Cells(fila, 11).Value = .dato_observaciones.Text
    
End With

End Sub


'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub copiar_tz14_datos_fijos(ByVal fila As Integer)

With userForm_tz14

    .TextBox_documento.Text = Cells(fila, 2).Value
    .TextBox_beneficiario.Text = Cells(fila, 3).Value & " " & Cells(fila, 4).Value
    .TextBox_fecha_obito.Text = Cells(fila, 5).Value
    .TextBox_fecha_comite.Text = Cells(fila, 6).Value
    
End With

End Sub


'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_tz14_copiar_datos_relevamiento(ByVal fila As Integer)

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz14
    
    'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
    'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
    'lineas de codigo de mas

    .dato_fecha_comite_pregunta.Text = Cells(fila, 8).Value
    If (.dato_fecha_comite_pregunta.Text = "") Then
        .dato_fecha_comite_pregunta.Text = leyenda
    End If
    
    .dato_fecha_comite_terreno.Text = Cells(fila, 9).Value
    If (.dato_fecha_comite_terreno.Text = "") Then
        .dato_fecha_comite_terreno.Text = leyenda
    End If
    
    .dato_diagnostico.Text = Cells(fila, 10).Value
    If (.dato_diagnostico.Text = "") Then
        .dato_diagnostico.Text = leyenda
    End If
    
    .dato_observaciones.Text = Cells(fila, 11).Value
    
End With

End Sub





'brief desbloquea y limpia los comboboxs y textboxs que corresponde a los datos obligatorios
'param void
'return void

Function userForm_tz14_permitir_campos_requeridos()

Dim leyenda As String
Dim leyenda2 As String
Dim leyenda3 As String

leyenda3 = "Dato no obligatorio"

    
With userForm_tz14
                    
    With .dato_fecha_comite_pregunta
        .Locked = False
        If (.Text = leyenda3) Then
            .Text = ""
        End If
        .BackColor = RGB(255, 255, 255)
    End With
    
    If (.dato_fecha_comite_pregunta = "No" Or .dato_fecha_comite_pregunta = "no" Or .dato_fecha_comite_pregunta = "") Then
        With .dato_fecha_comite_terreno
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
    End If
        
    With .dato_diagnostico
        .Locked = False
        If (.Text = leyenda3) Then
            .Text = ""
        End If
        .BackColor = RGB(255, 255, 255)
    End With

End With

End Function

