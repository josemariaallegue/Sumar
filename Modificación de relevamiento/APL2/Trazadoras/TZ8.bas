Attribute VB_Name = "TZ8"
'declaracion de variables globales

Public filaDobleClickTz8 As Integer

Public columnaDobleClickTz8 As Integer

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteTz8 As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorTz8 As Integer


'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_tz8_bloquear()

With userForm_tz8

    .TextBox_n_efector.Locked = True
    .TextBox_denominacion_efector.Locked = True
    .TextBox_beneficiario.Locked = True
    .TextBox_documento.Locked = True
    .TextBox_fecha_nacimiento.Locked = True
    .dato_fecha_vacuna_bacteriana_quintuple.Locked = True
    .dato_fecha_vacuna_antipoliomielitica.Locked = True
    
    .dato_vacuna_bacteriana_quintuple_pregunta.Locked = True
    .dato_vacuna_bacteriana_quintuple_terreno.Locked = True
    .dato_vacuna_antipoliomielitica_pregunta.Locked = True
    .dato_vacuna_antipoliomielitica_terreno.Locked = True
    

End With

End Function


'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_tz8_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz8
    
    .dato_vacuna_bacteriana_quintuple_pregunta.Text = leyenda
    .dato_vacuna_bacteriana_quintuple_terreno.Text = leyenda
    .dato_vacuna_antipoliomielitica_pregunta.Text = leyenda
    .dato_vacuna_antipoliomielitica_terreno.Text = leyenda
    
    .dato_vacuna_bacteriana_quintuple_pregunta.BackColor = RGB(169, 169, 169)
    .dato_vacuna_bacteriana_quintuple_terreno.BackColor = RGB(169, 169, 169)
    .dato_vacuna_antipoliomielitica_pregunta.BackColor = RGB(169, 169, 169)
    .dato_vacuna_antipoliomielitica_terreno.BackColor = RGB(169, 169, 169)
    
    .dato_vacuna_bacteriana_quintuple_pregunta.Locked = True
    .dato_vacuna_bacteriana_quintuple_terreno.Locked = True
    .dato_vacuna_antipoliomielitica_pregunta.Locked = True
    .dato_vacuna_antipoliomielitica_terreno.Locked = True
    
End With

End Function


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_tz8_verificacion_blancos() As Integer

With userForm_tz8

    If (.dato_fuente.Text = "" Or .dato_fecha_vacuna_bacteriana_quintuple.Text = "" Or .dato_vacuna_bacteriana_quintuple_pregunta.Text = "" Or _
    .dato_vacuna_bacteriana_quintuple_terreno.Text = "" Or .dato_fecha_vacuna_antipoliomielitica.Text = "" Or .dato_vacuna_antipoliomielitica_pregunta.Text = "" Or _
    .dato_vacuna_antipoliomielitica_terreno.Text = "") Then
        
        userForm_tz8_verificacion_blancos = 1

    Else

        userForm_tz8_verificacion_blancos = 0
    
    End If
 
End With

End Function




'brief: copia los datos del userform_tz8 al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_tz8_guardar_datos(ByVal fila As Integer)

With userForm_tz8
    
    Cells(fila, 10).Value = .dato_fuente.Text
    Cells(fila, 13).Value = .dato_vacuna_bacteriana_quintuple_pregunta.Text
    Cells(fila, 14).Value = .dato_vacuna_bacteriana_quintuple_terreno.Text
    Cells(fila, 16).Value = .dato_vacuna_antipoliomielitica_pregunta.Text
    Cells(fila, 17).Value = .dato_vacuna_antipoliomielitica_terreno.Text
    Cells(fila, 18).Value = .dato_observaciones.Text
    
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

Sub copiar_tz8_datos_fijos(ByVal fila As Integer)

With userForm_tz8

    .TextBox_n_efector.Text = Cells(fila, 3).Value
    .TextBox_denominacion_efector.Text = Cells(fila, 4).Value
    .TextBox_documento.Text = Cells(fila, 5).Value
    .TextBox_beneficiario.Text = Cells(fila, 6).Value & " " & Cells(fila, 7).Value
    .TextBox_fecha_nacimiento.Text = Cells(fila, 8).Value
    .dato_fecha_vacuna_bacteriana_quintuple.Text = Cells(fila, 12).Value
    .dato_fecha_vacuna_antipoliomielitica.Text = Cells(fila, 15).Value
    
End With

End Sub


'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_tz8_copiar_datos_relevamiento(ByVal fila As Integer)

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_tz8
    
    .dato_fuente.Text = Cells(fila, 10).Value
    
    'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
    If (.dato_fuente.Text <> "No consta fuente de información" And .dato_fuente.Text <> "Prestación inexistente") Then
        
        'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
        'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
        'lineas de codigo de mas
        .dato_vacuna_bacteriana_quintuple_pregunta.Text = Cells(fila, 13).Value
        If (.dato_vacuna_bacteriana_quintuple_pregunta.Text = "") Then
            .dato_vacuna_bacteriana_quintuple_pregunta.Text = leyenda
        End If
        
        .dato_vacuna_bacteriana_quintuple_terreno.Text = Cells(fila, 14).Value
        If (.dato_vacuna_bacteriana_quintuple_terreno.Text = "") Then
            .dato_vacuna_bacteriana_quintuple_terreno.Text = leyenda
        End If
        
        .dato_vacuna_antipoliomielitica_pregunta.Text = Cells(fila, 16).Value
        If (.dato_vacuna_antipoliomielitica_pregunta.Text = "") Then
            .dato_vacuna_antipoliomielitica_pregunta.Text = leyenda
        End If
        
        .dato_vacuna_antipoliomielitica_terreno.Text = Cells(fila, 17).Value
        If (.dato_vacuna_antipoliomielitica_terreno.Text = "") Then
            .dato_vacuna_antipoliomielitica_terreno.Text = leyenda
        End If
        
    Else
    
        Call userform_tz8_dato_no_obligatorio
        
    End If
    
    .dato_observaciones.Text = Cells(fila, 18).Value
    
End With

End Sub



'brief desbloquea y limpia los comboboxs y textboxs que corresponde a los datos obligatorios
'param void
'return void

Function userForm_tz8_permitir_campos_requeridos()

Dim leyenda As String
Dim leyenda2 As String
Dim leyenda3 As String

leyenda = "Labrar acta"
leyenda2 = "Labrar acta e indicar fuente de información en observaciones"
leyenda3 = "Dato no obligatorio"

'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_tz8.dato_validacion.Text <> leyenda And userForm_tz8.dato_validacion.Text <> leyenda2) Then
    
    With userForm_tz8
                    
        With .dato_vacuna_bacteriana_quintuple_pregunta
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        If (.dato_vacuna_bacteriana_quintuple_terreno = "No" Or .dato_vacuna_bacteriana_quintuple_terreno = "no" Or .dato_vacuna_bacteriana_quintuple_terreno = "") Then
            With .dato_vacuna_bacteriana_quintuple_terreno
                .Locked = False
                If (.Text = leyenda3) Then
                    .Text = ""
                End If
                .BackColor = RGB(255, 255, 255)
            End With
        End If
        
        With .dato_vacuna_antipoliomielitica_pregunta
            .Locked = False
            If (.Text = leyenda3) Then
                .Text = ""
            End If
            .BackColor = RGB(255, 255, 255)
        End With
        
        If (.dato_vacuna_antipoliomielitica_terreno = "No" Or .dato_vacuna_antipoliomielitica_terreno = "no" Or .dato_vacuna_antipoliomielitica_terreno = "") Then
            With .dato_vacuna_antipoliomielitica_terreno
                .Locked = False
                If (.Text = leyenda3) Then
                    .Text = ""
                End If
                .BackColor = RGB(255, 255, 255)
            End With
        End If
        
    End With
    
End If

End Function
