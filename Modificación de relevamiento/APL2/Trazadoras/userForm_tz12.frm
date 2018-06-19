VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_tz12 
   Caption         =   "Formulario de relevamiento - Trazadora 12"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14715
   OleObjectBlob   =   "userForm_tz12.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_tz12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_tz12_verificacion_blancos = 1) Then
    MsgBox ("No se han completado todos los campos")
End If


If (auxiliarInexistenteTz12 = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

    auxiliarInexistenteTz12 = 0

End If


Call userForm_tz12_guardar_datos(filaDobleClicktz12)

MsgBox ("Se han guardado con exito")


If (userForm_tz12.dato_validacion.Text = "Labrar acta" Or userForm_tz12.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClicktz12, columnaDobleClickTz12).Value = "Labrar acta"

ElseIf (userForm_tz12_verificacion_blancos <> 1) Then

    Cells(filaDobleClicktz12, columnaDobleClickTz12).Value = "Completo"

Else

    Cells(filaDobleClicktz12, columnaDobleClickTz12).Value = "Incompleto"

End If

Unload Me

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

Private Sub dato_firma_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_firma.Text <> "Dato no obligatorio" And dato_firma.Text <> "Si" And dato_firma.Text <> "si" And dato_firma.Text <> "No" _
And dato_firma.Text <> "no") Then

dato_firma.Text = ""

End If

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_tz12.dato_observaciones.MultiLine = True

End Sub

Private Sub dato_reporte_histologico_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_reporte_histologico.Text <> "1 = H-SIL" And dato_reporte_histologico.Text <> "2 = CIN 2" And dato_reporte_histologico.Text <> "3 = CIN 3" And _
dato_reporte_histologico.Text <> "4 = Carcinoma in situ" And dato_reporte_histologico.Text <> "5 = Cáncer cervico uterino" And dato_reporte_histologico.Text <> "No consta" And _
dato_reporte_histologico.Text <> "Dato no obligatorio") Then

dato_reporte_histologico.Text = ""

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
auxiliarInexistenteTz12 = 0

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "SITAM" Or dato_fuente.Text = "LAP" Or dato_fuente = "HC" Or _
dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente") Then


    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)
        Call userForm_tz12_permitir_campos_requeridos
        auxiliarInexistenteTz12 = 0
        
    Else
        'con estos 3 if le doy valor y formato al textbox de verificacion de casos
        'ademas se bloquean todas los campos y se pone "Dato no obligatorio"
        If (dato_fuente.Text = "No consta fuente de información") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz12_dato_no_obligatorio

        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz12_dato_no_obligatorio
            auxiliarInexistenteTz12 = 1


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
.AddItem "LAP"
.AddItem "HC"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
End With

'consulta sobre la fecha de diagnostico encontrada
With dato_fecha_diagnostico_pregunta
.AddItem "Si"
.AddItem "No"
End With

'reporte histologico
With dato_reporte_histologico
.AddItem "1 = H-SIL"
.AddItem "2 = CIN 2"
.AddItem "3 = CIN 3"
.AddItem "4 = Carcinoma in situ"
.AddItem "5 = Cáncer cervico uterino"
.AddItem "No consta"
End With

'consulta sobre la fecha de tratamiento encontrada
With dato_fecha_tratamiento_pregunta
.AddItem "Si"
.AddItem "No"
End With

'firma
With dato_firma
.AddItem "Si"
.AddItem "No"
End With


''Verifica si hay un valor en la celda de fuente de informacion y si es falso pone la leyenda
''"Ingresar fuente de informacion"
'If (Cells(filaDobleClickTZ12, columnaDobleClickTZ12 + 1) = "") Then
'dato_validacion.Text = "Ingresar fuente de información"
'dato_validacion.BackColor = RGB(255, 255, 0)
'End If


Application.EnableEvents = True

End Sub



