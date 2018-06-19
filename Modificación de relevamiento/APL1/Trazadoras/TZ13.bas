Attribute VB_Name = "TZ13"
'declaracion de variables globales

Public filaDobleClickTz13 As Integer

Public columnaDobleClickTz13 As Integer

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteTz13 As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorTz13 As Integer


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_tz13_bloquear()

With userForm_tz13

    .TextBox_n_efector.Locked = True
    .TextBox_denominacion_efector.Locked = True
    .TextBox_beneficiario.Locked = True
    .TextBox_documento.Locked = True
    .TextBox_fecha_nacimiento.Locked = True
    
    .dato_fecha_tratamiento.Locked = True
    .dato_fecha_tratamiento_pregunta.Locked = True
    .dato_fecha_tratamiento_terreno.Locked = True
    .dato_fecha_diagnostico.Locked = True
    .dato_fecha_diagnostico_pregunta.Locked = True
    .dato_fecha_diagnostico_terreno.Locked = True
    .dato_diagnostico.Locked = True
    .dato_tamaño.Locked = True
    .dato_ganglios.Locked = True
    .dato_metastasis.Locked = True
    .dato_estadio.Locked = True
    
    
End With

End Function


'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_tz13_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz13
    
    .dato_fecha_tratamiento_pregunta.Text = leyenda
    .dato_fecha_tratamiento_terreno.Text = leyenda
    .dato_fecha_diagnostico_pregunta.Text = leyenda
    .dato_fecha_diagnostico_terreno.Text = leyenda
    .dato_diagnostico.Text = leyenda
    .dato_tamaño.Text = leyenda
    .dato_ganglios.Text = leyenda
    .dato_metastasis.Text = leyenda
    .dato_estadio.Text = leyenda
    
    .dato_fecha_tratamiento_pregunta.BackColor = RGB(169, 169, 169)
    .dato_fecha_tratamiento_terreno.BackColor = RGB(169, 169, 169)
    .dato_fecha_diagnostico_pregunta.BackColor = RGB(169, 169, 169)
    .dato_fecha_diagnostico_terreno.BackColor = RGB(169, 169, 169)
    .dato_diagnostico.BackColor = RGB(169, 169, 169)
    .dato_tamaño.BackColor = RGB(169, 169, 169)
    .dato_ganglios.BackColor = RGB(169, 169, 169)
    .dato_metastasis.BackColor = RGB(169, 169, 169)
    .dato_estadio.BackColor = RGB(169, 169, 169)

    .dato_fecha_tratamiento_pregunta.Locked = True
    .dato_fecha_tratamiento_terreno.Locked = True
    .dato_fecha_diagnostico_pregunta.Locked = True
    .dato_fecha_diagnostico_terreno.Locked = True
    .dato_diagnostico.Locked = True
    .dato_tamaño.Locked = True
    .dato_ganglios.Locked = True
    .dato_metastasis.Locked = True
    .dato_estadio.Locked = True
    
End With

End Function


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_tz13_verificacion_blancos() As Integer

With userForm_tz13

    If (.dato_fuente.Text = "" Or .dato_fecha_tratamiento.Text = "" Or .dato_fecha_tratamiento_pregunta.Text = "" Or _
    .dato_fecha_tratamiento_terreno.Text = "" Or .dato_fecha_diagnostico.Text = "" Or .dato_fecha_diagnostico_pregunta.Text = "" Or _
    .dato_fecha_diagnostico_terreno.Text = "" Or .dato_diagnostico.Text = "" Or .dato_tamaño.Text = "" Or .dato_ganglios.Text = "" Or _
    .dato_metastasis.Text = "" Or .dato_estadio.Text = "") Then
        
        userForm_tz13_verificacion_blancos = 1

    Else

        userForm_tz13_verificacion_blancos = 0
    
    End If
 
End With

End Function




'brief: copia los datos del userform_tz13 al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_tz13_guardar_datos(ByVal fila As Integer)

With userForm_tz13
    
    Cells(fila, 10).Value = .dato_fuente.Text
    Cells(fila, 13).Value = .dato_fecha_tratamiento_pregunta.Text
    Cells(fila, 14).Value = .dato_fecha_tratamiento_terreno.Text
    Cells(fila, 16).Value = .dato_fecha_diagnostico_pregunta.Text
    Cells(fila, 17).Value = .dato_fecha_diagnostico_terreno.Text
    Cells(fila, 18).Value = .dato_diagnostico.Text
    Cells(fila, 19).Value = .dato_tamaño.Text
    Cells(fila, 20).Value = .dato_ganglios.Text
    Cells(fila, 21).Value = .dato_metastasis.Text
    Cells(fila, 22).Value = .dato_estadio.Text
    Cells(fila, 23).Value = .dato_observaciones.Text
    
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

Sub copiar_tz13_datos_fijos(ByVal fila As Integer)

With userForm_tz13

    .TextBox_n_efector.Text = Cells(fila, 3).Value
    .TextBox_denominacion_efector.Text = Cells(fila, 4).Value
    .TextBox_documento.Text = Cells(fila, 5).Value
    .TextBox_beneficiario.Text = Cells(fila, 6).Value & " " & Cells(fila, 7).Value
    .TextBox_fecha_nacimiento.Text = Cells(fila, 8).Value
    .dato_fecha_tratamiento.Text = Cells(fila, 12).Value
    .dato_fecha_diagnostico.Text = Cells(fila, 15).Value
    
End With

End Sub


'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_tz13_copiar_datos_relevamiento(ByVal fila As Integer)

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz13
    
    .dato_fuente.Text = Cells(fila, 10).Value
    
    'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
    If (.dato_fuente.Text <> "No consta fuente de información" And .dato_fuente.Text <> "Prestación inexistente") Then
    
        'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
        'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
        'lineas de codigo de mas
        .dato_fecha_tratamiento_pregunta.Text = Cells(fila, 13).Value
        If (.dato_fecha_tratamiento_pregunta.Text = "") Then
            .dato_fecha_tratamiento_pregunta.Text = leyenda
        End If
        
        .dato_fecha_tratamiento_terreno.Text = Cells(fila, 14).Value
        If (.dato_fecha_tratamiento_terreno.Text = "") Then
            .dato_fecha_tratamiento_terreno.Text = leyenda
        End If
        
        .dato_fecha_diagnostico_pregunta.Text = Cells(fila, 16).Value
        If (.dato_fecha_diagnostico_pregunta.Text = "") Then
            .dato_fecha_diagnostico_pregunta.Text = leyenda
        End If
        
        .dato_fecha_diagnostico_terreno.Text = Cells(fila, 17).Value
        If (.dato_fecha_diagnostico_terreno.Text = "") Then
            .dato_fecha_diagnostico_terreno.Text = leyenda
        End If
        
        .dato_diagnostico.Text = Cells(fila, 18).Value
        If (.dato_diagnostico.Text = "") Then
            .dato_diagnostico.Text = leyenda
        End If
        
        .dato_tamaño.Text = Cells(fila, 19).Value
        If (.dato_tamaño.Text = "") Then
            .dato_tamaño.Text = leyenda
        End If
        
        .dato_ganglios.Text = Cells(fila, 20).Value
        If (.dato_ganglios.Text = "") Then
            .dato_ganglios.Text = leyenda
        End If
        
        .dato_metastasis.Text = Cells(fila, 21).Value
        If (.dato_metastasis.Text = "") Then
            .dato_metastasis.Text = leyenda
        End If
        
        .dato_estadio.Text = Cells(fila, 22).Value
        If (.dato_estadio.Text = "") Then
            .dato_estadio.Text = leyenda
        End If
        
    Else
        
        Call userform_tz13_dato_no_obligatorio
    
    End If
    
    .dato_observaciones.Text = Cells(fila, 23).Value
    
End With

End Sub



'brief desbloquea y limpia los comboboxs y textboxs que corresponde a los datos obligatorios
'param void
'return void

Function userForm_tz13_permitir_campos_requeridos()

Dim leyenda As String
Dim leyenda2 As String
Dim leyenda3 As String

leyenda = "Labrar acta"
leyenda2 = "Labrar acta e indicar fuente de información en observaciones"
leyenda3 = "Dato no obligatorio"

'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_tz13.dato_validacion.Text <> leyenda And userForm_tz13.dato_validacion.Text <> leyenda2) Then
    
    With userForm_tz13
                    
        With .dato_fecha_tratamiento_pregunta
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        If (.dato_fecha_tratamiento_pregunta = "No" Or .dato_fecha_tratamiento_pregunta = "no" Or .dato_fecha_tratamiento_pregunta = "") Then
            With .dato_fecha_tratamiento_terreno
                .Locked = False
                If (.Text = leyenda3) Then
                    .Text = ""
                End If
                .BackColor = RGB(255, 255, 255)
            End With
        End If
        
        With .dato_fecha_diagnostico_pregunta
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        If (.dato_fecha_diagnostico_pregunta = "No" Or .dato_fecha_diagnostico_pregunta = "no" Or .dato_fecha_diagnostico_pregunta = "") Then
            With .dato_fecha_diagnostico_terreno
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
        
        With .dato_tamaño
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_ganglios
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_metastasis
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        With .dato_estadio
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
          
    End With
    
End If

End Function
