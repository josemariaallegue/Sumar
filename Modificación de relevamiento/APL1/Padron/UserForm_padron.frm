VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_padron 
   Caption         =   "Formulario de relevamiento"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15345
   OleObjectBlob   =   "UserForm_padron.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm_padron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'brief: cierra el userForm_padron
'param: void
'retrurn: void

Private Sub CommandButton1_Click()

Dim respuesta As String
Dim fuenteInformacion As String


If (auxiliarInexistente = 1 And dato_observaciones.Text = "") Then

fuenteInformacion = InputBox("Por favor ingrese la fuente de información", "Fuente de información")

dato_observaciones.Text = fuenteInformacion

auxiliarInexistente = 0

End If



If (auxiliarGuardado = 0) Then

    respuesta = MsgBox("No se ha guardado. ¿Desea guardar antes de salir?", vbYesNo)
        
    If (respuesta = vbYes) Then
        
    Call CommandButton3_Click
        
    Unload Me
    
    Else
    
    Unload Me
    
    End If
    
ElseIf (auxiliarGuardado = 1) Then
    
    respuesta = MsgBox("Se han realizado cambios. ¿Desea guardar antes de salir?", vbYesNo)
    
    If (respuesta = vbYes) Then
        
    Call CommandButton3_Click
        
    Unload Me
    
    Else
    
    Unload Me
    
    End If
    
Else
    
    respuesta = MsgBox("¿Esta seguro que desea salir?", vbYesNo)
        
    If (respuesta = vbYes) Then
        
    Unload Me
    
    Else
    
    End If
    
End If



End Sub

'brief: guarda los cambios
'param: void
'retrurn: void
Public Sub CommandButton3_Click()

Dim respuesta As String


'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_padron_verificacion_blancos = 1) Then
MsgBox ("No se han completado todos los campos")
End If


'consulta si se deseaguardar los datos y si es verdadero, copia los datos
'al formulario y pone completo si se completaron todos las celdas e incompleto si no
'tambien cambia el valor de auxiliarGuardar para poder comprobar si se guardo antes de salir
respuesta = MsgBox("¿Esta seguro que desea guardar?", vbYesNo)

If (respuesta = vbYes) Then

    Call userForm_padron_guardar_datos(filaDobleClick)
    
    MsgBox ("Se han guardado con exito")
    
    If (userForm_padron_verificacion_blancos <> 1) Then
    Cells(filaDobleClick, columnaDobleClick).Value = "Completo"
    
    Else
    Cells(filaDobleClick, columnaDobleClick).Value = "Incompleto"
    End If
    
    auxiliarGuardado = 2

Else

MsgBox ("No se ha guardado")

End If

End Sub

Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_padron_verificacion_blancos = 1) Then

    MsgBox ("No se han completado todos los campos")
    
End If


If (auxiliarInexistente = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad.", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

auxiliarInexistente = 0

End If


Call userForm_padron_guardar_datos(filaDobleClick)
    
MsgBox ("Se han guardado con exito")

    
If (dato_validacion.Text = "Labrar acta" Or dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClick, columnaDobleClick).Value = "Labrar acta"
    
ElseIf (dato_validacion.Text = "Caso duplicado") Then

    Cells(filaDobleClick, columnaDobleClick).Value = "Caso duplicado"

    
ElseIf (userForm_padron_verificacion_blancos <> 1) Then

    Cells(filaDobleClick, columnaDobleClick).Value = "Completo"

Else

    Cells(filaDobleClick, columnaDobleClick).Value = "Incompleto"
    
End If

Unload Me

End Sub

Private Sub dato_fuente_Change()

Dim flag As Integer
Dim concatenacion As String
Dim resultadoConsultaV As Variant
Dim parte As String
Dim fuente As String

parte = Left(TextBox_codigo.Text, 3)
fuente = UserForm_padron.dato_fuente.Text


'resetea el valor de esta variable por si el auditor se confunde al colocar la fuente.
'si se pone primero prestacion inexistente y luego no consta fuente de informacion sigue pidiendo que se ingrese la fuente, con esto no
auxiliarInexistente = 0


'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "FM" Or dato_fuente.Text = "HC" Or dato_fuente.Text = "HCPB" Or dato_fuente.Text = "FOD" _
Or dato_fuente.Text = "LE" Or dato_fuente.Text = "EPICRISIS" Or dato_fuente.Text = "LL" Or dato_fuente.Text = "REGAP" _
Or dato_fuente.Text = "LSI" Or dato_fuente.Text = "PGRUP" Or dato_fuente.Text = "SI" Or dato_fuente.Text = "RV" _
Or dato_fuente.Text = "SIP" Or dato_fuente.Text = "SITAM" Or dato_fuente.Text = "R.PAIERC-SISA" Or dato_fuente.Text = "HCORP" Or dato_fuente.Text = "PAS" _
Or dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente" Or dato_fuente.Text = "Caso duplicado") Then
    
    
    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)
        
        auxiliarInexistente = 0
        
        '(Jose: no recuerdo por que estan pero por precaucion no modificar)
        'este if evita que haya problema con el inializador de valores del combobox y que haya error la
        'primera vez que se ejeuta el programa
        If (codigoDobleClick <> "") Then
            
            'si se modifica un valor de este campo vuelvo a revisar los campos obligatorio
            'es para que si se pone "no consta fuente de informacion" y luego se cambia la fuente no quede todo bloqueado
            Call userForm_padron_permitir_campos_requeridos(codigoDobleClick)
        
        End If
            
        
        'formulas para textbox de control de fuente de informacion
        'se realizan concatenaciones para los consultaV
        'como los consultaV pueden devolver error 1004 (no se encuentra el dato buscado) se deben hacer if
        On Error Resume Next
        
            concatenacion = TextBox_codigo.Text & dato_fuente.Text & ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 31).Value
            
            resultadoConsultaV = Application.WorksheetFunction.VLookup(concatenacion, ThisWorkbook.Sheets("Fuentes de informacion validas").Range("F1:F1100"), 1, False)
            
            'si se encuentra el valor de una la fuente es valida
            If (Err.Number <> 1004) Then
            
                dato_control_fuente.Text = "Fuente valida"
                dato_control_fuente.BackColor = RGB(87, 166, 57)
            
            'si no, verificanos que la prestacion sea de embarazo y luego hacemos otro consultaV
            'y verificamos que no devuelva error
            
            ElseIf ((parte = "PRP" And (fuente = "HC" Or fuente = "HCPB" Or fuente = "FM" Or fuente = "FOD" Or fuente = "HCORP" Or fuente = "LL")) _
            Or (parte = "LBL" And (fuente = "HC" Or fuente = "HCPB" Or fuente = "LL")) _
            Or parte = "IGR" And (fuente = "FM" Or fuente = "HCPB" Or fuente = "LSI" Or fuente = "HC" Or fuente = "SITAM")) Then
            
            Call userForm_padron_permitir_campos_requeridos(codigoDobleClick)
            
            dato_control_fuente.Text = "Fuente valida"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            
            Else
            
                resultadoConsultaV = Application.WorksheetFunction.VLookup(TextBox_codigo.Text, ThisWorkbook.Sheets("Fuentes de informacion validas").Range("B1:D1100"), 3, False)

                    
                If (resultadoConsultaV = "Embarazo") Then
                   
                    'se debe hacer otro "on error" porque si no se toma el valor del anterior
                    On Error Resume Next
                        
                        concatenacion = TextBox_codigo.Text & dato_fuente.Text
                        
                        resultadoConsultaV = Application.WorksheetFunction.VLookup(concatenacion, ThisWorkbook.Sheets("Fuentes de informacion validas").Range("E1:E1100"), 1, False)
                                
                                
                        If (Err.Number <> 1004) Then
                        
                            dato_control_fuente.Text = "Fuente valida"
                            dato_control_fuente.BackColor = RGB(87, 166, 57)
                                
                            
                        Else
                            
                            dato_control_fuente.Text = "Fuente invalida"
                            dato_control_fuente.BackColor = RGB(255, 0, 0)
                            
                            dato_validacion.Text = "Labrar acta"
                            dato_validacion.BackColor = RGB(255, 0, 0)
                            
                            Call userForm_padron_dato_no_obligatorio
                            
                        End If
                            
                    On Error GoTo 0
                            
                        
                Else
                    
                    dato_control_fuente.Text = "Fuente invalida"
                    dato_control_fuente.BackColor = RGB(255, 0, 0)
                    
                    dato_validacion.Text = "Labrar acta"
                    dato_validacion.BackColor = RGB(255, 0, 0)
                    
                    Call userForm_padron_dato_no_obligatorio
                    
                End If
                
            End If
        
        
        On Error GoTo 0
    
    
    Else
    
        'con estos 3 if le doy valor y formato al textbox de verificacion de casos
        'ademas se bloquean todas los campos y se pone "Dato no obligatorio"
        If (dato_fuente.Text = "No consta fuente de información") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userForm_padron_dato_no_obligatorio
                
        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userForm_padron_dato_no_obligatorio
            auxiliarInexistente = 1
            
            
        ElseIf (dato_fuente.Text = "Caso duplicado") Then
            dato_validacion.Text = "Caso duplicado"
            dato_validacion.BackColor = RGB(255, 160, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userForm_padron_dato_no_obligatorio
            
            
                
        ElseIf (dato_fuente.Text = "") Then
            dato_validacion.Text = "Ingresar la fuente de información"
            dato_validacion.BackColor = RGB(255, 255, 0)
            
        Else
            dato_validacion.Text = "Ok"
            dato_validacion.BackColor = RGB(255, 255, 255)
             
        End If
        
    End If
    
Else
    
dato_fuente.Text = ""
    
End If

'para evitar la modificacion del "control de fuente de informacion" y "verificacion de casos"
dato_control_fuente.Locked = True
dato_validacion.Locked = True

auxiliarGuardado = 1
    
End Sub

Private Sub dato_control_prenatal_Change()

If Not (IsNumeric(dato_control_prenatal.Text) Or dato_control_prenatal = "Dato no obligatorio") Then

dato_control_prenatal.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_diagnostico_Change()

'para que siga escribiendo en una linea de inferior
dato_diagnostico.MultiLine = True

auxiliarGuardado = 1

End Sub


Private Sub dato_estudios_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_estudios.Text <> "Evaluación genitourinaria y examen mamario" And dato_estudios.Text <> "Evalución genitourinaria y colonoscopia" _
And dato_estudios.Text <> "Odontograma" And dato_estudios.Text <> "Medición de agudeza visual" And dato_estudios.Text <> "Evaluación genitourinaria" _
And dato_estudios.Text <> "Examen mamario" And dato_estudios.Text <> "Colonoscopia" And dato_estudios.Text <> "No consta" And dato_estudios.Text <> "No requiere" _
And dato_estudios.Text <> "Dato no obligatorio") Then
    
dato_estudios.Text = ""
    
End If

auxiliarGuardado = 1


End Sub

Private Sub dato_evaluacion_riesgo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_evaluacion_riesgo.Text <> "Dato no obligatorio" And dato_evaluacion_riesgo.Text <> "Si" And dato_evaluacion_riesgo.Text <> "si" And dato_evaluacion_riesgo.Text <> "No" _
And dato_evaluacion_riesgo.Text <> "no") Then

dato_evaluacion_riesgo.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_ta_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_ta.Text <> "Dato no obligatorio" And dato_ta.Text <> "Si" And dato_ta.Text <> "si" And dato_ta.Text <> "No" _
And dato_ta.Text <> "no") Then

dato_ta.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_imc_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_imc.Text <> "Dato no obligatorio" And dato_imc.Text <> "Si" And dato_imc.Text <> "si" And dato_imc.Text <> "No" _
And dato_imc.Text <> "no") Then

dato_imc.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_percentilo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_percentilo.Text <> "Dato no obligatorio" And dato_percentilo.Text <> "Si" And dato_percentilo.Text <> "si" And dato_percentilo.Text <> "No" _
And dato_percentilo.Text <> "no") Then

dato_percentilo.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_peso_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_peso.Text <> "Dato no obligatorio" And dato_peso.Text <> "Si" And dato_peso.Text <> "si" And dato_peso.Text <> "No" _
And dato_peso.Text <> "no") Then

dato_peso.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_talla_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_talla.Text <> "Dato no obligatorio" And dato_talla.Text <> "Si" And dato_talla.Text <> "si" And dato_talla.Text <> "No" _
And dato_talla.Text <> "no") Then

dato_talla.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_plan_seguimiento_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_plan_seguimiento.Text <> "Dato no obligatorio" And dato_plan_seguimiento.Text <> "Si" And dato_plan_seguimiento.Text <> "si" And dato_plan_seguimiento.Text <> "No" _
And dato_plan_seguimiento.Text <> "no" And dato_plan_seguimiento.Text <> "No requiere") Then

dato_plan_seguimiento.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_tratamiento_instaurado_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_tratamiento_instaurado.Text <> "Dato no obligatorio" And dato_tratamiento_instaurado.Text <> "Si" And dato_tratamiento_instaurado.Text <> "si" And dato_tratamiento_instaurado.Text <> "No" _
And dato_tratamiento_instaurado.Text <> "no" And dato_tratamiento_instaurado.Text <> "No requiere") Then

dato_tratamiento_instaurado.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_transcripcion_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_transcripcion.Text <> "Dato no obligatorio" And dato_transcripcion.Text <> "Si" And dato_transcripcion.Text <> "si" And dato_transcripcion.Text <> "No" _
And dato_transcripcion.Text <> "no" And dato_transcripcion.Text <> "No requiere") Then

dato_transcripcion.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_constancia_inmunizaciones_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_constancia_inmunizaciones.Text <> "Dato no obligatorio" And dato_constancia_inmunizaciones.Text <> "Si" And dato_constancia_inmunizaciones.Text <> "si" And dato_constancia_inmunizaciones.Text <> "No" _
And dato_constancia_inmunizaciones.Text <> "no") Then

dato_constancia_inmunizaciones.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_firma_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_firma.Text <> "Dato no obligatorio" And dato_firma.Text <> "Si" And dato_firma.Text <> "si" And dato_firma.Text <> "No" _
And dato_firma.Text <> "no") Then

dato_firma.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_sello_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_sello.Text <> "Dato no obligatorio" And dato_sello.Text <> "Si" And dato_sello.Text <> "si" And dato_sello.Text <> "No" _
And dato_sello.Text <> "no") Then

dato_sello.Text = ""

End If

auxiliarGuardado = 1

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
UserForm_padron.dato_observaciones.MultiLine = True

auxiliarGuardado = 1


End Sub

Private Sub dato_control_fuente_Change()

auxiliarGuardado = 1

End Sub

Private Sub dato_validacion_Change()

'para que siga escribiendo en una linea de inferior
dato_validacion.MultiLine = True

auxiliarGuardado = 1

End Sub


Private Sub TextBox_beneficiario_Change()

'para que siga escribiendo en una linea de inferior
TextBox_beneficiario.MultiLine = True

End Sub

Private Sub TextBox_denominacion_efector_Change()

'para que siga escribiendo en una linea de inferior
TextBox_denominacion_efector.MultiLine = True

End Sub

Private Sub TextBox_documento_Change()

'para darle formato de documento
TextBox_documento = Format(TextBox_documento, "##,##00")

End Sub

Private Sub TextBox_edad_Change()

'para darle formato de numero con 2 decimales
TextBox_edad.Text = Format(TextBox_edad.Text, "0.00")

End Sub

Private Sub TextBox_linea_cuidado_Change()

'para que siga escribiendo en una linea de inferior
UserForm_padron.TextBox_linea_cuidado.MultiLine = True

End Sub

Private Sub userform_initialize()

Application.EnableEvents = False


CommandButton4.Caption = "Guardar" & Chr(10) & "y salir"



'otorgo valores a los comboboxs
With dato_fuente
.AddItem "FM"
.AddItem "HC"
.AddItem "HCPB"
.AddItem "FOD"
.AddItem "LE"
.AddItem "EPICRISIS"
.AddItem "LL"
.AddItem "REGAP"
.AddItem "LSI"
.AddItem "PGRUP"
.AddItem "SI"
.AddItem "RV"
.AddItem "SIP"
.AddItem "SITAM"
.AddItem "R.PAIERC-SISA"
.AddItem "HCORP"
.AddItem "PAS"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
.AddItem "Caso duplicado"
End With


With dato_estudios
.AddItem "Evaluación genitourinaria y examen mamario"
.AddItem "Evalución genitourinaria y colonoscopia"
.AddItem "Odontograma"
.AddItem "Medición de agudeza visual"
.AddItem "Evaluación genitourinaria"
.AddItem "Examen mamario"
.AddItem "Colonoscopia"
.AddItem "No consta"
.AddItem "No requiere"
End With

With dato_evaluacion_riesgo
.AddItem "Si"
.AddItem "No"
End With

With dato_ta
.AddItem "Si"
.AddItem "No"
End With

With dato_imc
.AddItem "Si"
.AddItem "No"
End With

With dato_percentilo
.AddItem "Si"
.AddItem "No"
End With

With dato_peso
.AddItem "Si"
.AddItem "No"
End With

With dato_talla
.AddItem "Si"
.AddItem "No"
End With

With dato_plan_seguimiento
.AddItem "Si"
.AddItem "No"
.AddItem "No requiere"
End With

With dato_tratamiento_instaurado
.AddItem "Si"
.AddItem "No"
.AddItem "No requiere"
End With

With dato_transcripcion
.AddItem "Si"
.AddItem "No"
.AddItem "No requiere"
End With

With dato_constancia_inmunizaciones
.AddItem "Si"
.AddItem "No"
End With

With dato_firma
.AddItem "Si"
.AddItem "No"
End With

With dato_sello
.AddItem "Si"
.AddItem "No"
End With


'Verifica si hay un valor en la celda de fuente de informacion y si es falso pone la leyenda
'"Ingresar fuente de informacion"
If (Cells(filaDobleClick, columnaDobleClick + 1) = "") Then
dato_validacion.Text = "Ingresar fuente de información"
dato_validacion.BackColor = RGB(255, 255, 0)
End If


Application.EnableEvents = True



End Sub
