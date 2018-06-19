Attribute VB_Name = "niños_en_internacion"
'declaracion de variables globales

Public filaDobleClickNI As Integer

Public columnaDobleClickNI As Integer

Public codigoDobleClickNI As String

'existe para obligar al auditor a poner la fuente de informacion
'0 es falso (la prestacion no existe)
'1 es verdadero (la prestacion existe)
Public auxiliarInexistenteNI As Integer

'no se reconoce un codigo valido y saltea la apertura del userform
Public errorNI As Integer


'brief: copia los datos del beneficiario del formulario en el userform
'param: recibe el rango donde se hizo doble click
'return: void

Sub copiar_ni_datos_fijos(ByVal Target As Range)

With userForm_ni

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
    .TextBox_monto = Cells(Target.Row, Target.Column + 10).Value
    .TextBox_edad.Text = Cells(Target.Row, Target.Column + 12).Value

End With

'verifica que si la celda de nombre de prestacion contiene "La prestacion no corresponde al grupo poblacional" si
'esto es verdadero le otorga un valor al textbox de control de fuente de informacion
If (Cells(filaDobleClickNI, columnaDobleClickNI - 2).Value = "La prestación no corresponde al grupo poblacional") Then

    userForm_ni.dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"

End If

End Sub

'brief: bloquea la los textboxs y comboboxs del userform
'param: void
'return: void

Function userForm_ni_bloquear()

With userForm_ni

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
    
    .dato_transcripcion_estudios.Locked = True
    .dato_tratamiento_instaurado.Locked = True
    .dato_firma.Locked = True
    .dato_sello.Locked = True
    .dato_contrarreferencia.Locked = True

End With

End Function

'brief: copia los datos ya relevados a un userForm
'param: rango donde se hizo doble click
'return: void

Sub userForm_ni_copiar_datos_relevamiento(ByVal Target As Range)


Dim leyenda As String

leyenda = "Dato no obligatorio"


userForm_ni.dato_fuente.Text = Cells(Target.Row, Target.Column + 1).Value

'este primer if es por si el auditor modifica la fuente de informacion fuera del userfom
If (userForm_ni.dato_fuente.Text <> "No consta fuente de información" And userForm_ni.dato_fuente.Text <> "Prestación inexistente" And _
userForm_ni.dato_control_fuente.Text <> "Fuente invalida") Then

    userForm_ni.dato_diagnostico.Text = Cells(Target.Row, Target.Column + 5).Value
    
    'los siguientes if verifican si la celda donde esta el valor esta vacia, si lo esta le ponen al textbox correspondiente
    'el valor de leyenda. Esto se hace porque la mayoria de las prestaciones tienen pocos datos para relevar y esto evita
    'lineas de codigo de mas
    userForm_ni.dato_transcripcion_estudios.Text = Cells(Target.Row, Target.Column + 4).Value
    If (userForm_ni.dato_transcripcion_estudios.Text = "") Then
        userForm_ni.dato_transcripcion_estudios.Text = leyenda
    End If
    
    userForm_ni.dato_tratamiento_instaurado.Text = Cells(Target.Row, Target.Column + 6).Value
    If (userForm_ni.dato_tratamiento_instaurado.Text = "") Then
        userForm_ni.dato_tratamiento_instaurado.Text = leyenda
    End If
    
    userForm_ni.dato_contrarreferencia.Text = Cells(Target.Row, Target.Column + 7).Value
    If (userForm_ni.dato_contrarreferencia.Text = "") Then
        userForm_ni.dato_contrarreferencia.Text = leyenda
    End If
    
    userForm_ni.dato_firma.Text = Cells(Target.Row, Target.Column + 8).Value
    If (userForm_ni.dato_firma.Text = "") Then
        userForm_ni.dato_firma.Text = leyenda
    End If
    
    userForm_ni.dato_sello.Text = Cells(Target.Row, Target.Column + 9).Value
    If (userForm_ni.dato_sello.Text = "") Then
        userForm_ni.dato_sello.Text = leyenda
    End If
    
    'userForm_ni.dato_control_fuente.Text = Cells(target.Row, target.Column + 2).Value
    
    'userForm_ni.dato_validacion.Text = Cells(target.Row, target.Column + 21).Value
    
Else
    
    Call userform_ni_dato_no_obligatorio

End If

userForm_ni.dato_observaciones.Text = Cells(Target.Row, Target.Column + 11).Value

End Sub


'brief: pone el valor de leyenda en los textbos y comboboxs correspondientes, pinta las celdas de color gris y las bloquea
'param: void
'return: void

Function userform_ni_dato_no_obligatorio()

Dim leyenda As String

leyenda = "Dato no obligatorio"

With userForm_ni
    
    .dato_diagnostico.Text = leyenda
    .dato_transcripcion_estudios.Text = leyenda
    .dato_tratamiento_instaurado.Text = leyenda
    .dato_firma.Text = leyenda
    .dato_sello.Text = leyenda
    .dato_contrarreferencia.Text = leyenda
    
    .dato_diagnostico.BackColor = RGB(169, 169, 169)
    .dato_transcripcion_estudios.BackColor = RGB(169, 169, 169)
    .dato_tratamiento_instaurado.BackColor = RGB(169, 169, 169)
    .dato_firma.BackColor = RGB(169, 169, 169)
    .dato_sello.BackColor = RGB(169, 169, 169)
    .dato_contrarreferencia.BackColor = RGB(169, 169, 169)
    
    .dato_diagnostico.Locked = True
    .dato_transcripcion_estudios.Locked = True
    .dato_tratamiento_instaurado.Locked = True
    .dato_firma.Locked = True
    .dato_sello.Locked = True
    .dato_contrarreferencia.Locked = True
    
End With

End Function


'brief: copia los datos del userform_ni al formulario
'param: es la fila donde se hizo doble click
'return: void

Sub userForm_ni_guardar_datos(ByVal fila As Integer)

With userForm_ni
    
    Cells(fila, 14).Value = .dato_fuente.Text
    Cells(fila, 17).Value = .dato_transcripcion_estudios.Text
    Cells(fila, 18).Value = .dato_diagnostico.Text
    Cells(fila, 19).Value = .dato_tratamiento_instaurado.Text
    Cells(fila, 20).Value = .dato_contrarreferencia.Text
    Cells(fila, 21).Value = .dato_firma.Text
    Cells(fila, 22).Value = .dato_sello.Text
    Cells(fila, 24).Value = .dato_observaciones.Text
    
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

Sub userForm_ni_permitir_campos_requeridos(ByVal codigo As String)

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
If (userForm_ni.dato_validacion.Text <> leyenda And userForm_ni.dato_validacion.Text <> leyenda2 And _
userForm_ni.dato_validacion.Text <> leyenda3) Then

    If (userForm_ni.dato_diagnostico.Text = leyenda4) Then
    
    userForm_ni.dato_diagnostico.Locked = False
    userForm_ni.dato_diagnostico.BackColor = RGB(255, 255, 255)
    userForm_ni.dato_diagnostico.Text = ""
    
    End If
    
End If


'este if evita que se haga el for al pedo si no consta fuente de informacion o la prestacion es inexistente
If (userForm_ni.dato_validacion.Text <> leyenda And userForm_ni.dato_validacion.Text <> leyenda2 And userForm_ni.dato_validacion.Text <> leyenda3) Then


    For i = 1 To 250
    
        'coincidencia de poblacion con el codigo o el codigo con la poblacion embarazo
        If ((ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 23).Value And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo) Or (ThisWorkbook.Sheets("Requerimientos").Cells(i, 4).Value = codigo And _
        ThisWorkbook.Sheets("Requerimientos").Cells(i, 1).Value = "Embarazos")) Then
                
                've si el codigo es CTC001A97 y la poblacion es niños (distintos datos obligatorios por edad)
                If (codigo = "CTC001A97" And ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 23).Value = "Niños") Then
                    
                    'hace coincidir la edad en el relevamiento con la descripcion
                    If ((ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 12).Value < 1 And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "menores de 1") Or _
                    ((ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 12).Value >= 1 And ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 12).Value < 6) And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "1 a 5") Or _
                    (ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 12).Value >= 6 And ThisWorkbook.Sheets("Requerimientos").Cells(i, 3).Value = "mayores a 6")) Then

                        For j = 5 To 31
                            
                            'verifica que que el dato que corresponde a la celda se obligatorio
                            If (ThisWorkbook.Sheets("Requerimientos").Cells(i, j).Value <> "") Then
                                
                                'copia el nombre del dato obligatorio y empiesa a verificar cual es
                                texto = ThisWorkbook.Sheets("Requerimientos").Cells(1, j).Value
                                
                                Select Case texto
                                
                                    Case "Informe o transcripción de estudios solicitados"
                                    
                                        With userForm_ni.dato_transcripcion_estudios
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Tratamiento instaurado"
                                    
                                        With userForm_ni.dato_tratamiento_instaurado
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Contrarreferencia  o epicrisis de datos referidos al diagnostico y tratamiento indicado"
                                        
                                        With userForm_ni.dato_contrarreferencia
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Firma"
                                    
                                        With userForm_ni.dato_firma
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Sello"
                                        
                                        With userForm_ni.dato_sello
                                        .Locked = False
                                        If (.Text = leyenda4) Then
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
                                
                                Case "Informe o transcripción de estudios solicitados"
                                    
                                        With userForm_ni.dato_transcripcion_estudios
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Tratamiento instaurado"
                                    
                                        With userForm_ni.dato_tratamiento_instaurado
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Contrarreferencia  o epicrisis de datos referidos al diagnostico y tratamiento indicado"
                                        
                                        With userForm_ni.dato_contrarreferencia
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Firma"
                                    
                                        With userForm_ni.dato_firma
                                        .Locked = False
                                        If (.Text = leyenda4) Then
                                            .Text = ""
                                        End If
                                        .BackColor = RGB(255, 255, 255)
                                        End With
                                        
                                        
                                    Case "Sello"
                                        
                                        With userForm_ni.dato_sello
                                        .Locked = False
                                        If (.Text = leyenda4) Then
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


'brief: verifica si alguno de los campos fue dejado en blanco
'param: void
'return: 1 si alguna de las celdas esta vacia
'        0 si estan todas completas

Function userForm_ni_verificacion_blancos() As Integer

With userForm_ni

    If (.dato_fuente.Text = "" Or .dato_diagnostico = "" Or .dato_transcripcion_estudios = "" Or .dato_tratamiento_instaurado = "" Or _
    .dato_contrarreferencia = "" Or .dato_firma = "" Or .dato_sello = "") Then
        
        userForm_ni_verificacion_blancos = 1

    Else

        userForm_ni_verificacion_blancos = 0
    
    End If
 
End With

End Function
