VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_tz13 
   Caption         =   "Formulario de relevamiento - Trazadora 13"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17505
   OleObjectBlob   =   "userForm_tz13.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_tz13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_tz13_verificacion_blancos = 1) Then
    MsgBox ("No se han completado todos los campos")
End If


If (auxiliarInexistenteTz13 = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

    auxiliarInexistenteTz13 = 0

End If


Call userForm_tz13_guardar_datos(filaDobleClickTz13)

MsgBox ("Se han guardado con exito")


If (userForm_tz13.dato_validacion.Text = "Labrar acta" Or userForm_tz13.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClickTz13, columnaDobleClickTz13).Value = "Labrar acta"

ElseIf (userForm_tz13_verificacion_blancos <> 1) Then

    Cells(filaDobleClickTz13, columnaDobleClickTz13).Value = "Completo"

Else

    Cells(filaDobleClickTz13, columnaDobleClickTz13).Value = "Incompleto"

End If

Unload Me

End Sub



Private Sub dato_amenorrea_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_amenorrea.Text <> "Dato no obligatorio" And dato_amenorrea.Text <> "Si" And dato_amenorrea.Text <> "si" And dato_amenorrea.Text <> "No" _
And dato_amenorrea.Text <> "no") Then

dato_amenorrea.Text = ""

End If

End Sub

Private Sub dato_diagnostico_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_diagnostico.Text <> "1 = Carcinoma in situ" And dato_diagnostico.Text <> "2 = Carcinoma invasor" And _
dato_diagnostico.Text <> "No consta" And dato_diagnostico.Text <> "Dato no obligatorio") Then

dato_diagnostico.Text = ""

End If

End Sub

Private Sub dato_estadio_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_estadio.Text <> "I" And dato_estadio.Text <> "IIA" And dato_estadio.Text <> "IIB" And dato_estadio.Text <> "IIIA" And _
dato_estadio.Text <> "IIIB" And dato_estadio.Text <> "IIIC" And dato_estadio.Text <> "IV" And dato_estadio.Text <> "No consta" _
And dato_estadio.Text <> "Dato no obligatorio") Then

    dato_estadio.Text = ""

End If

End Sub

Private Sub dato_fecha_diagnostico_pregunta_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fecha_diagnostico_pregunta.Text <> "Dato no obligatorio" And dato_fecha_diagnostico_pregunta.Text <> "Si" And dato_fecha_diagnostico_pregunta.Text <> "si" And dato_fecha_diagnostico_pregunta.Text <> "No" _
And dato_fecha_diagnostico_pregunta.Text <> "no") Then

dato_fecha_diagnostico_pregunta.Text = ""

End If


'bloquea y desbloquea el campos de diagnostico encontrado en terreno dependiendo del valor ingresado
If (dato_fecha_diagnostico_pregunta.Text = "Si" Or dato_fecha_diagnostico_pregunta.Text = "si") Then

    With dato_fecha_diagnostico_terreno
        .Text = "Dato no obligatorio"
        .BackColor = RGB(169, 169, 169)
        .Locked = True
    End With

ElseIf (dato_fecha_diagnostico_pregunta.Text = "No" Or dato_fecha_diagnostico_pregunta.Text = "no") Then
    
    With dato_fecha_diagnostico_terreno
        .Locked = False
        If (.Text = "Dato no obligatorio") Then
            .Text = ""
        End If
        .BackColor = RGB(255, 255, 255)
    End With

End If

End Sub

Private Sub dato_fecha_diagnostico_terreno_Change()

End Sub

Private Sub dato_fecha_tratamiento_pregunta_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fecha_tratamiento_pregunta.Text <> "Dato no obligatorio" And dato_fecha_tratamiento_pregunta.Text <> "Si" And dato_fecha_tratamiento_pregunta.Text <> "si" And dato_fecha_tratamiento_pregunta.Text <> "No" _
And dato_fecha_tratamiento_pregunta.Text <> "no") Then

dato_fecha_tratamiento_pregunta.Text = ""

End If


'bloquea y desbloquea el campos de tratamiento encontrado en terreno dependiendo del valor ingresado
If (dato_fecha_tratamiento_pregunta.Text = "Si" Or dato_fecha_tratamiento_pregunta.Text = "si") Then

    With dato_fecha_tratamiento_terreno
        .Text = "Dato no obligatorio"
        .BackColor = RGB(169, 169, 169)
        .Locked = True
    End With

ElseIf (dato_fecha_tratamiento_pregunta.Text = "No" Or dato_fecha_tratamiento_pregunta.Text = "no") Then
    
    With dato_fecha_tratamiento_terreno
        .Locked = False
        If (.Text = "Dato no obligatorio") Then
            .Text = ""
        End If
        .BackColor = RGB(255, 255, 255)
    End With

End If

End Sub


Private Sub dato_fecha_tratamiento_terreno_Change()

End Sub

Private Sub dato_ganglios_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_ganglios.Text <> "N0" And dato_ganglios.Text <> "N1" And dato_ganglios.Text <> "N2" And dato_ganglios.Text <> "No consta" And _
dato_ganglios.Text <> "Dato no obligatorio") Then

dato_ganglios.Text = ""

End If

End Sub

Private Sub dato_metastasis_Change()

If (dato_metastasis.Text <> "M0" And dato_metastasis.Text <> "M1" And dato_metastasis <> "No consta" And _
dato_metastasis.Text <> "Dato no obligatorio") Then

dato_metastasis.Text = ""

End If

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_tz13.dato_observaciones.MultiLine = True

End Sub


Private Sub dato_tamaño_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_tamaño.Text <> "T0" And dato_tamaño.Text <> "T1" And dato_tamaño.Text <> "T2" And dato_tamaño.Text <> "T3" And _
dato_tamaño.Text <> "T4" And dato_tamaño.Text <> "No consta" And dato_tamaño.Text <> "Dato no obligatorio") Then

dato_tamaño.Text = ""

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
auxiliarInexistenteTz13 = 0

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "SITAM" Or dato_fuente.Text = "RITA" Or dato_fuente = "HC" Or dato_fuente = "RAP" Or _
dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente") Then


    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)
        Call userForm_tz13_permitir_campos_requeridos
        auxiliarInexistenteTz13 = 0
        
    Else
        'con estos 3 if le doy valor y formato al textbox de verificacion de casos
        'ademas se bloquean todas los campos y se pone "Dato no obligatorio"
        If (dato_fuente.Text = "No consta fuente de información") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz13_dato_no_obligatorio

        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz13_dato_no_obligatorio
            auxiliarInexistenteTz13 = 1


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

Private Sub TextBox_edad_Change()

'para darle formato de numero con 2 decimales
TextBox_edad.Text = Format(TextBox_edad.Text, "0.00")

End Sub


Private Sub userform_initialize()

Application.EnableEvents = False


CommandButton4.Caption = "Guardar" & Chr(10) & "y salir"


'otorgo valores a los comboboxs

'fuente de informacion
With dato_fuente
.AddItem "SITAM"
.AddItem "RITA"
.AddItem "HC"
.AddItem "RAP"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
End With

'consulta sobre la fecha de diagnostico encontrada
With dato_fecha_diagnostico_pregunta
.AddItem "Si"
.AddItem "No"
End With

'consulta sobre la fecha de tratamiento encontrada
With dato_fecha_tratamiento_pregunta
.AddItem "Si"
.AddItem "No"
End With

'diagnostico
With dato_diagnostico
.AddItem "1 = Carcinoma in situ"
.AddItem "2 = Carcinoma invasor"
.AddItem "No consta"
End With

'tamaño
With dato_tamaño
.AddItem "T0"
.AddItem "T1"
.AddItem "T2"
.AddItem "T3"
.AddItem "T4"
.AddItem "No consta"
End With

'ganglios linfaticos
With dato_ganglios
.AddItem "N0"
.AddItem "N1"
.AddItem "N2"
.AddItem "No consta"
End With

'metastasis
With dato_metastasis
.AddItem "M0"
.AddItem "M1"
.AddItem "No consta"
End With

'estadio
With dato_estadio
.AddItem "I"
.AddItem "IIA"
.AddItem "IIB"
.AddItem "IIIA"
.AddItem "IIIB"
.AddItem "IIIC"
.AddItem "IV"
.AddItem "No consta"
End With


''Verifica si hay un valor en la celda de fuente de informacion y si es falso pone la leyenda
''"Ingresar fuente de informacion"
'If (Cells(filaDobleClicktz13, columnaDobleClicktz13 + 1) = "") Then
'dato_validacion.Text = "Ingresar fuente de información"
'dato_validacion.BackColor = RGB(255, 255, 0)
'End If


Application.EnableEvents = True

End Sub

