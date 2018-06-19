VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_ni 
   Caption         =   "Formulario de relevamiento - Niños en internación"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15435
   OleObjectBlob   =   "userForm_ni.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_ni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_ni_verificacion_blancos = 1) Then

    MsgBox ("No se han completado todos los campos")
    
End If


If (auxiliarInexistenteNI = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad.", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
        
    End If

auxiliarInexistenteNI = 0

End If


Call userForm_ni_guardar_datos(filaDobleClickNI)

MsgBox ("Se han guardado con exito")


If (userForm_ni.dato_validacion.Text = "Labrar acta" Or userForm_ni.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClickNI, columnaDobleClickNI).Value = "Labrar acta"

ElseIf (userForm_ni_verificacion_blancos <> 1) Then

    Cells(filaDobleClickNI, columnaDobleClickNI).Value = "Completo"

Else

    Cells(filaDobleClickNI, columnaDobleClickNI).Value = "Incompleto"

End If

Unload Me

End Sub

Private Sub dato_contrarreferencia_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_contrarreferencia.Text <> "Dato no obligatorio" And dato_contrarreferencia.Text <> "Si" And dato_contrarreferencia.Text <> "si" And dato_contrarreferencia.Text <> "No" _
And dato_contrarreferencia.Text <> "no") Then

dato_contrarreferencia.Text = ""

End If

End Sub

Private Sub dato_diagnostico_Change()

'para que siga escribiendo en una linea de inferior
dato_diagnostico.MultiLine = True

End Sub


Private Sub dato_transcripcion_estudios_Change()
    
'con este if se evita que se ingresen datos que no estan permitidos
If (dato_transcripcion_estudios.Text <> "Dato no obligatorio" And dato_transcripcion_estudios.Text <> "Si" And dato_transcripcion_estudios.Text <> "si" And dato_transcripcion_estudios.Text <> "No" _
And dato_transcripcion_estudios.Text <> "no" And dato_transcripcion_estudios.Text <> "No requiere") Then

dato_transcripcion_estudios.Text = ""

End If

End Sub

Private Sub dato_tratamiento_instaurado_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_tratamiento_instaurado.Text <> "Dato no obligatorio" And dato_tratamiento_instaurado.Text <> "Si" And dato_tratamiento_instaurado.Text <> "si" And dato_tratamiento_instaurado.Text <> "No" _
And dato_tratamiento_instaurado.Text <> "no" And dato_tratamiento_instaurado.Text <> "No requiere") Then

dato_tratamiento_instaurado.Text = ""

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
userForm_ni.dato_observaciones.MultiLine = True

End Sub


Private Sub dato_validacion_Change()

'para que siga escribiendo en una linea de inferior
dato_validacion.MultiLine = True

End Sub

Private Sub dato_fuente_Change()


'resetea el valor de esta variable por si el auditor se confunde al colocar la fuente.
'si se pone primero prestacion inexistente y luego no consta fuente de informacion sigue pidiendo que se ingrese la fuente, con esto no
auxiliarInexistenteNI = 0

Dim flag As Integer
Dim concatenacion As String
Dim resultadoConsultaV As Variant

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "FM" Or dato_fuente.Text = "HC" Or dato_fuente.Text = "HCPB" Or dato_fuente.Text = "FOD" _
Or dato_fuente.Text = "LE" Or dato_fuente.Text = "EPICRISIS" Or dato_fuente.Text = "LL" Or dato_fuente.Text = "REGAP" _
Or dato_fuente.Text = "LSI" Or dato_fuente.Text = "PGRUP" Or dato_fuente.Text = "SI" Or dato_fuente.Text = "RV" _
Or dato_fuente.Text = "SIP" Or dato_fuente.Text = "SITAM" Or dato_fuente.Text = "No consta fuente de información" _
Or dato_fuente.Text = "Prestación inexistente" Or dato_fuente.Text = "Caso duplicado") Then


    If (Cells(filaDobleClickNI, columnaDobleClickNI - 2).Value = "La prestación no corresponde al grupo poblacional") Then

        dato_control_fuente.Text = "La prestación no corresponde al grupo poblacional"

    End If

    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)

        auxiliarInexistenteNI = 0

        '(Jose: no recuerdo por que estan pero por precaucion no modificar)
        'este if evita que haya problema con el inializador de valores del combobox y que haya error la
        'primera vez que se ejeuta el programa
        If (codigoDobleClickNI <> "") Then

            'si se modifica un valor de este campo vuelvo a revisar los campos obligatorio
            'es para que si se pone "no consta fuente de informacion" y luego se cambia la fuente no quede todo bloqueado
            Call userForm_ni_permitir_campos_requeridos(codigoDobleClickNI)

        End If


        'formulas para textbox de control de fuente de informacion
        'se realizan concatenaciones para los consultaV
        'como los consultaV pueden devolver error 1004 (no se encuentra el dato buscado) se deben hacer if
        On Error Resume Next

            concatenacion = TextBox_codigo.Text & dato_fuente.Text & ActiveSheet.Cells(filaDobleClickNI, columnaDobleClickNI + 23).Value
            
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

                            Call userform_ni_dato_no_obligatorio

                        End If

                    On Error GoTo 0


                Else

                    dato_control_fuente.Text = "Fuente invalida"
                    dato_control_fuente.BackColor = RGB(255, 0, 0)

                    dato_validacion.Text = "Labrar acta"
                    dato_validacion.BackColor = RGB(255, 0, 0)

                    Call userform_ni_dato_no_obligatorio

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
            Call userform_ni_dato_no_obligatorio

        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userform_ni_dato_no_obligatorio
            auxiliarInexistenteNI = 1


        ElseIf (dato_fuente.Text = "Caso duplicado") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            dato_control_fuente.Text = "N/A"
            dato_control_fuente.BackColor = RGB(87, 166, 57)
            Call userform_ni_dato_no_obligatorio



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


Private Sub dato_vida_fetal_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_vida_fetal.Text <> "Dato no obligatorio" And dato_vida_fetal.Text <> "Si" And dato_vida_fetal.Text <> "si" And dato_vida_fetal.Text <> "No" _
And dato_vida_fetal.Text <> "no") Then

dato_vida_fetal.Text = ""

End If

End Sub



Private Sub Frame3_Click()

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
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
End With


'transcripcion de esutdios
With dato_transcripcion_estudios
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

'contrarreferencia o epicrisis de datos referidos
'al diagnostico y tratamiento indicado
With dato_contrarreferencia
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
If (Cells(filaDobleClickNI, columnaDobleClickNI + 1) = "") Then
dato_validacion.Text = "Ingresar la fuente de información"
dato_validacion.BackColor = RGB(255, 255, 0)
End If


Application.EnableEvents = True
'
'
'
End Sub

