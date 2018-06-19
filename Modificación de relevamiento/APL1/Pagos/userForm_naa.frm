VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_naa 
   Caption         =   "Formulario de relevamiento - Niños, adolescentes y adultos"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15345
   OleObjectBlob   =   "userForm_naa.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_naa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_naa_verificacion_blancos = 1) Then

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


Call userForm_naa_guardar_datos(filaDobleClick)
    
MsgBox ("Se han guardado con exito")


If (userForm_naa.dato_validacion.Text = "Labrar acta" Or userForm_naa.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClick, columnaDobleClick).Value = "Labrar acta"
    
ElseIf (userForm_naa_verificacion_blancos <> 1) Then

    Cells(filaDobleClick, columnaDobleClick).Value = "Completo"

Else

    Cells(filaDobleClick, columnaDobleClick).Value = "Incompleto"
    
End If

Unload Me

End Sub
Private Sub dato_diagnostico_Change()

'para que siga escribiendo en una linea de inferior
dato_diagnostico.MultiLine = True

End Sub


Private Sub dato_estudios_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_estudios.Text <> "Evaluación genitourinaria y examen mamario" And dato_estudios.Text <> "Evalución genitourinaria y colonoscopia" _
And dato_estudios.Text <> "Odontograma" And dato_estudios.Text <> "Medición de agudeza visual" And dato_estudios.Text <> "Evaluación genitourinaria" _
And dato_estudios.Text <> "Examen mamario" And dato_estudios.Text <> "Colonoscopia" And dato_estudios.Text <> "No consta" And dato_estudios.Text <> "No requiere" _
And dato_estudios.Text <> "Dato no obligatorio") Then
    
dato_estudios.Text = ""
    
End If

End Sub

Private Sub dato_evaluacion_riesgo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_evaluacion_riesgo.Text <> "Dato no obligatorio" And dato_evaluacion_riesgo.Text <> "Si" And dato_evaluacion_riesgo.Text <> "si" And dato_evaluacion_riesgo.Text <> "No" _
And dato_evaluacion_riesgo.Text <> "no") Then

dato_evaluacion_riesgo.Text = ""

End If

End Sub

Private Sub dato_fecha_notificacion_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fecha_notificacion.Text <> "Dato no obligatorio" And dato_fecha_notificacion.Text <> "Si" And dato_fecha_notificacion.Text <> "si" And dato_fecha_notificacion.Text <> "No" _
And dato_fecha_notificacion.Text <> "no") Then

dato_fecha_notificacion.Text = ""

End If

End Sub

Private Sub dato_ta_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_ta.Text <> "Dato no obligatorio" And dato_ta.Text <> "Si" And dato_ta.Text <> "si" And dato_ta.Text <> "No" _
And dato_ta.Text <> "no") Then

dato_ta.Text = ""

End If

End Sub

Private Sub dato_imc_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_imc.Text <> "Dato no obligatorio" And dato_imc.Text <> "Si" And dato_imc.Text <> "si" And dato_imc.Text <> "No" _
And dato_imc.Text <> "no") Then

dato_imc.Text = ""

End If

End Sub

Private Sub dato_percentilo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_percentilo.Text <> "Dato no obligatorio" And dato_percentilo.Text <> "Si" And dato_percentilo.Text <> "si" And dato_percentilo.Text <> "No" _
And dato_percentilo.Text <> "no") Then

dato_percentilo.Text = ""

End If

End Sub

Private Sub dato_peso_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_peso.Text <> "Dato no obligatorio" And dato_peso.Text <> "Si" And dato_peso.Text <> "si" And dato_peso.Text <> "No" _
And dato_peso.Text <> "no") Then

dato_peso.Text = ""

End If

End Sub

Private Sub dato_talla_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_talla.Text <> "Dato no obligatorio" And dato_talla.Text <> "Si" And dato_talla.Text <> "si" And dato_talla.Text <> "No" _
And dato_talla.Text <> "no") Then

dato_talla.Text = ""

End If

End Sub

Private Sub dato_plan_seguimiento_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_plan_seguimiento.Text <> "Dato no obligatorio" And dato_plan_seguimiento.Text <> "Si" And dato_plan_seguimiento.Text <> "si" And dato_plan_seguimiento.Text <> "No" _
And dato_plan_seguimiento.Text <> "no" And dato_plan_seguimiento.Text <> "No requiere") Then

dato_plan_seguimiento.Text = ""

End If

End Sub

Private Sub dato_tratamiento_instaurado_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_tratamiento_instaurado.Text <> "Dato no obligatorio" And dato_tratamiento_instaurado.Text <> "Si" And dato_tratamiento_instaurado.Text <> "si" And dato_tratamiento_instaurado.Text <> "No" _
And dato_tratamiento_instaurado.Text <> "no" And dato_tratamiento_instaurado.Text <> "No requiere") Then

dato_tratamiento_instaurado.Text = ""

End If

End Sub

Private Sub dato_transcripcion_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_transcripcion.Text <> "Dato no obligatorio" And dato_transcripcion.Text <> "Si" And dato_transcripcion.Text <> "si" And dato_transcripcion.Text <> "No" _
And dato_transcripcion.Text <> "no" And dato_transcripcion.Text <> "No requiere") Then

dato_transcripcion.Text = ""

End If

End Sub

Private Sub dato_constancia_inmunizaciones_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_constancia_inmunizaciones.Text <> "Dato no obligatorio" And dato_constancia_inmunizaciones.Text <> "Si" And dato_constancia_inmunizaciones.Text <> "si" And dato_constancia_inmunizaciones.Text <> "No" _
And dato_constancia_inmunizaciones.Text <> "no") Then

dato_constancia_inmunizaciones.Text = ""

End If

End Sub

Private Sub dato_firma_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_firma.Text <> "Dato no obligatorio" And dato_firma.Text <> "Si" And dato_firma.Text <> "si" And dato_firma.Text <> "No" _
And dato_firma.Text <> "no") Then

dato_firma.Text = ""

End If

End Sub

Private Sub dato_sello_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_sello.Text <> "Dato no obligatorio" And dato_sello.Text <> "Si" And dato_sello.Text <> "si" And dato_sello.Text <> "No" _
And dato_sello.Text <> "no") Then

dato_sello.Text = ""

End If

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_naa.dato_observaciones.MultiLine = True

End Sub

Private Sub dato_validacion_Change()

'para que siga escribiendo en una linea de inferior
dato_validacion.MultiLine = True

End Sub

Private Sub dato_fuente_Change()


'resetea el valor de esta variable por si el auditor se confunde al colocar la fuente.
'si se pone primero prestacion inexistente y luego no consta fuente de informacion sigue pidiendo que se ingrese la fuente, con esto no
auxiliarInexistente = 0


Dim flag As Integer
Dim concatenacion As String
Dim resultadoConsultaV As Variant


'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "FM" Or dato_fuente.Text = "HC" Or dato_fuente.Text = "HCPB" Or dato_fuente.Text = "FOD" _
Or dato_fuente.Text = "LE" Or dato_fuente.Text = "EPICRISIS" Or dato_fuente.Text = "LL" Or dato_fuente.Text = "REGAP" _
Or dato_fuente.Text = "LSI" Or dato_fuente.Text = "PGRUP" Or dato_fuente.Text = "SI" Or dato_fuente.Text = "RV" _
Or dato_fuente.Text = "SIP" Or dato_fuente.Text = "SITAM" Or dato_fuente.Text = "LG" _
Or dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente" _
Or dato_fuente.Text = "Caso duplicado") Then
    
    If (Cells(filaDobleClick, columnaDobleClick - 2).Value = "La prestación no corresponde al grupo poblacional") Then
    
        dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"
        
    End If
    
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
            Call userForm_naa_permitir_campos_requeridos(codigoDobleClick)
        
        End If
            
        
        'formulas para textbox de control de fuente de informacion
        'se realizan concatenaciones para los consultaV
        'como los consultaV pueden devolver error 1004 (no se encuentra el dato buscado) se deben hacer if
        On Error Resume Next
        
            concatenacion = TextBox_codigo.Text & dato_fuente.Text & ActiveSheet.Cells(filaDobleClick, columnaDobleClick + 32).Value
            
            resultadoConsultaV = Application.WorksheetFunction.VLookup(concatenacion, ThisWorkbook.Sheets("Fuentes de informacion validas").Range("F1:F700"), 1, False)
            'si se encuentra el valor de una la fuente es valida
            If (Err.Number <> 1004) Then
            
                dato_control_fuente.Text = "Fuente valida"
                dato_control_fuente.BackColor = RGB(87, 166, 57)
            
            'si no, verificanos que la prestacion sea de embarazo y luego hacemos otro consultaV
            'y verificamos que no devuelva error
            Else
            
                resultadoConsultaV = Application.WorksheetFunction.VLookup(TextBox_codigo.Text, ThisWorkbook.Sheets("Fuentes de informacion validas").Range("B1:D700"), 3, False)

                    
                If (resultadoConsultaV = "Embarazo") Then
                   
                    'se debe hacer otro "on error" porque si no se toma el valor del anterior
                    On Error Resume Next
                        
                        concatenacion = TextBox_codigo.Text & dato_fuente.Text
                        
                        resultadoConsultaV = Application.WorksheetFunction.VLookup(concatenacion, ThisWorkbook.Sheets("Fuentes de informacion validas").Range("E1:E700"), 1, False)
                                
                                
                        If (Err.Number <> 1004) Then
                        
                            dato_control_fuente.Text = "Fuente valida"
                            dato_control_fuente.BackColor = RGB(87, 166, 57)
                                
                            
                        Else
                            
                            dato_control_fuente.Text = "Fuente invalida"
                            dato_control_fuente.BackColor = RGB(255, 0, 0)
                            
                            dato_validacion.Text = "Labrar acta"
                            dato_validacion.BackColor = RGB(255, 0, 0)
                            
                            Call userForm_naa_dato_no_obligatorio
                            
                        End If
                            
                    On Error GoTo 0
                            
                        
                Else
                    
                    dato_control_fuente.Text = "Fuente invalida"
                    dato_control_fuente.BackColor = RGB(255, 0, 0)
                    
                    dato_validacion.Text = "Labrar acta"
                    dato_validacion.BackColor = RGB(255, 0, 0)
                    
                    Call userForm_naa_dato_no_obligatorio
                    
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
            Call userForm_naa_dato_no_obligatorio
                
        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userForm_naa_dato_no_obligatorio
            auxiliarInexistente = 1
            
            
        ElseIf (dato_fuente.Text = "Caso duplicado") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userForm_naa_dato_no_obligatorio
            
            
                
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

End Sub




Private Sub TextBox_beneficiario_Change()

'para que siga escribiendo en una linea de inferior
TextBox_beneficiario.MultiLine = True

End Sub

Private Sub TextBox_denominacion_efector_Change()

'para que siga escribiendo en una linea de inferior
TextBox_denominacion_efector.MultiLine = True

End Sub

Private Sub TextBox_descripcion_Change()

'para que siga escribiendo en una linea de inferior
TextBox_descripcion.MultiLine = True

End Sub

Private Sub TextBox_documento_Change()

'para darle formato de documento
TextBox_documento = Format(TextBox_documento, "#,###,##")

End Sub

Private Sub TextBox_edad_Change()

'para darle formato de numero con 2 decimales
TextBox_edad.Text = Format(TextBox_edad.Text, "0.00")

End Sub


Private Sub TextBox_monto_Change()

TextBox_monto.Text = Format(TextBox_monto.Text, "Currency")


End Sub

Private Sub userform_initialize()

Application.EnableEvents = False


CommandButton4.Caption = "Guardar" & Chr(10) & "y salir"



'otorgo valores a los comboboxs

'fuente de informacion
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
.AddItem "LG"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
End With

'estudios realizados
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


'evaluacion de riesgo individual
With dato_evaluacion_riesgo
.AddItem "Si"
.AddItem "No"
End With

'toma de ta
With dato_ta
.AddItem "Si"
.AddItem "No"
End With

'IMC
With dato_imc
.AddItem "Si"
.AddItem "No"
End With

'percentilo
With dato_percentilo
.AddItem "Si"
.AddItem "No"
End With

'peso
With dato_peso
.AddItem "Si"
.AddItem "No"
End With

'talla
With dato_talla
.AddItem "Si"
.AddItem "No"
End With

'plan de seguimiento
With dato_plan_seguimiento
.AddItem "Si"
.AddItem "No"
.AddItem "No requiere"
End With

'tratamiento instaurado
With dato_tratamiento_instaurado
.AddItem "Si"
.AddItem "No"
.AddItem "No requiere"
End With

'informe o transcripcion de estudios solicitados
With dato_transcripcion
.AddItem "Si"
.AddItem "No"
.AddItem "No requiere"
End With

'constancia de aplicacion de inmunizaciones
With dato_constancia_inmunizaciones
.AddItem "Si"
.AddItem "No"
End With

'fecha de notificacion, embarazo o parto
With dato_fecha_notificacion
.AddItem "Si"
.AddItem "No"
End With

'firma y aclaracion
With dato_firma
.AddItem "Si"
.AddItem "No"
End With

'sello
With dato_sello
.AddItem "Si"
.AddItem "No"
End With


'Verifica si hay un valor en la celda de fuente de informacion y si es falso pone la leyenda
'"Ingresar fuente de informacion"
If (Cells(filaDobleClick, columnaDobleClick + 1) = "") Then
dato_validacion.Text = "Ingresar la fuente de información"
dato_validacion.BackColor = RGB(255, 255, 0)
End If


Application.EnableEvents = True



End Sub
