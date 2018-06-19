VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_tz7 
   Caption         =   "Formulario de relevamiento - Trazadora 7"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13305
   OleObjectBlob   =   "userForm_tz7.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_tz7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_tz7_verificacion_blancos = 1) Then
    MsgBox ("No se han completado todos los campos")
End If


If (auxiliarInexistenteTz7 = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

    auxiliarInexistenteTz7 = 0

End If


Call userForm_tz7_guardar_datos(filaDobleClickTz7)

MsgBox ("Se han guardado con exito")


If (userForm_tz7.dato_validacion.Text = "Labrar acta" Or userForm_tz7.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClickTz7, columnaDobleClickTz7).Value = "Labrar acta"

ElseIf (userForm_tz7_verificacion_blancos <> 1) Then

    Cells(filaDobleClickTz7, columnaDobleClickTz7).Value = "Completo"

Else

    Cells(filaDobleClickTz7, columnaDobleClickTz7).Value = "Incompleto"

End If

Unload Me

End Sub

Private Sub dato_fecha_tsomf_pregunta_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fecha_tsomf_pregunta.Text <> "Si" And dato_fecha_tsomf_pregunta.Text <> "si" And dato_fecha_tsomf_pregunta.Text <> "No" And dato_fecha_tsomf_pregunta.Text <> "no" And _
dato_fecha_tsomf_pregunta.Text <> "Dato no obligatorio") Then

dato_fecha_tsomf_pregunta.Text = ""

End If


'bloquea y desbloquea el campos de diagnostico encontrado en terreno dependiendo del valor ingresado
If (dato_fecha_tsomf_pregunta.Text = "Si" Or dato_fecha_tsomf_pregunta.Text = "si") Then

    With dato_fecha_tsmof_terreno
        .Text = "Dato no obligatorio"
        .BackColor = RGB(169, 169, 169)
        .Locked = True
    End With

ElseIf (dato_fecha_tsomf_pregunta.Text = "No" Or dato_fecha_tsomf_pregunta.Text = "no") Then
    
    With dato_fecha_tsmof_terreno
        .Locked = False
        If (.Text = "Dato no obligatorio") Then
            .Text = ""
        End If
        .BackColor = RGB(255, 255, 255)
    End With

End If

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_tz7.dato_observaciones.MultiLine = True

End Sub

Private Sub dato_resultado_tsomf_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_resultado_tsomf.Text <> "Positivo" And dato_resultado_tsomf.Text <> "Negativo" And dato_resultado_tsomf.Text <> "Dato no obligatorio") Then

dato_resultado_tsomf.Text = ""

End If
End Sub

Private Sub dato_validacion_Change()

'para que siga escribiendo en una linea de inferior
dato_validacion.MultiLine = True

End Sub

Private Sub dato_fuente_Change()

Dim flag As Integer
Dim concatenacion As String
Dim resultadoConsultaV As Variant

'resetea el valor de esta variable por si el auditor se confunde al colocar la fuente.
'si se pone primero prestacion inexistente y luego no consta fuente de informacion sigue pidiendo que se ingrese la fuente, con esto no
auxiliarInexistenteTz7 = 0

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "SITAM" Or dato_fuente.Text = "HC" Or dato_fuente.Text = "RL" Or dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente") Then


    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)
        Call userForm_tz7_permitir_campos_requeridos
        auxiliarInexistenteTz7 = 0

    Else
        'con estos 3 if le doy valor y formato al textbox de verificacion de casos
        'ademas se bloquean todas los campos y se pone "Dato no obligatorio"
        If (dato_fuente.Text = "No consta fuente de información") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz7_dato_no_obligatorio

        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz7_dato_no_obligatorio
            auxiliarInexistenteTz7 = 1


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

Private Sub TextBox_documento_Change()

'para darle formato de documento
TextBox_documento = Format(TextBox_documento, "#,###,##")

End Sub


Private Sub userform_initialize()

Application.EnableEvents = False


CommandButton4.Caption = "Guardar" & Chr(10) & "y salir"


'otorgo valores a los comboboxs

'fuente de informacion
With dato_fuente
.AddItem "SITAM"
.AddItem "HC"
.AddItem "RL"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
End With

'consulta sobre la fecha de tratamiento encontrada
With dato_fecha_tsomf_pregunta
.AddItem "Si"
.AddItem "No"
End With

'resultado del TSOMF
With dato_resultado_tsomf
.AddItem "Positivo"
.AddItem "Negativo"
End With

''Verifica si hay un valor en la celda de fuente de informacion y si es falso pone la leyenda
''"Ingresar fuente de informacion"
'If (Cells(filaDobleClickTZ7, columnaDobleClickTZ7 + 1) = "") Then
'dato_validacion.Text = "Ingresar fuente de información"
'dato_validacion.BackColor = RGB(255, 255, 0)
'End If


Application.EnableEvents = True

End Sub


