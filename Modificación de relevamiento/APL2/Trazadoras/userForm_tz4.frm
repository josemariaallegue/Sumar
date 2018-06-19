VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_tz4 
   Caption         =   "Formulario de relevamiento - Trazadora 4"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13230
   OleObjectBlob   =   "userForm_tz4.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_tz4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_tz4_verificacion_blancos = 1) Then
    MsgBox ("No se han completado todos los campos")
End If


If (auxiliarInexistenteTz4 = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

    auxiliarInexistenteTz4 = 0

End If


Call userForm_tz4_guardar_datos(filaDobleClickTz4)

MsgBox ("Se han guardado con exito")


If (userForm_tz4.dato_validacion.Text = "Labrar acta" Or userForm_tz4.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClickTz4, columnaDobleClickTz4).Value = "Labrar acta"

ElseIf (userForm_tz4_verificacion_blancos <> 1) Then

    Cells(filaDobleClickTz4, columnaDobleClickTz4).Value = "Completo"

Else

    Cells(filaDobleClickTz4, columnaDobleClickTz4).Value = "Incompleto"

End If

Unload Me

End Sub


Private Sub dato_firma_Change()

If (dato_firma.Text <> "Si" And dato_firma.Text <> "si" And dato_firma.Text <> "No" And dato_firma.Text <> "no" And _
dato_firma.Text <> "Dato no obligatorio") Then

dato_firma.Text = ""

End If

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_tz4.dato_observaciones.MultiLine = True

End Sub

Private Sub dato_perimetro_cefalico_Change()

If (dato_perimetro_cefalico.Text <> "Si" And dato_perimetro_cefalico.Text <> "si" And dato_perimetro_cefalico.Text <> "No" And dato_perimetro_cefalico.Text <> "no" And _
dato_perimetro_cefalico.Text <> "Dato no obligatorio") Then

dato_perimetro_cefalico.Text = ""

End If

End Sub

Private Sub dato_peso_Change()

If (dato_peso.Text <> "Si" And dato_peso.Text <> "si" And dato_peso.Text <> "No" And dato_peso.Text <> "no" And _
dato_peso.Text <> "Dato no obligatorio") Then

dato_peso.Text = ""

End If

End Sub

Private Sub dato_talla_Change()

If (dato_talla.Text <> "Si" And dato_talla.Text <> "si" And dato_talla.Text <> "No" And dato_talla.Text <> "no" And _
dato_talla.Text <> "Dato no obligatorio") Then

dato_talla.Text = ""

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
auxiliarInexistenteTz4 = 0

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "HC" Or dato_fuente.Text = "FM" Or dato_fuente.Text = "PP" Or dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente") Then


    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)
        Call userForm_tz4_permitir_campos_requeridos
        auxiliarInexistenteTz4 = 0
        
    Else
        'con estos 3 if le doy valor y formato al textbox de verificacion de casos
        'ademas se bloquean todas los campos y se pone "Dato no obligatorio"
        If (dato_fuente.Text = "No consta fuente de información") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz4_dato_no_obligatorio

        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz4_dato_no_obligatorio
            auxiliarInexistenteTz4 = 1


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
.AddItem "HC"
.AddItem "FM"
.AddItem "PP"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
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

'perimetro cefalico
With dato_perimetro_cefalico
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
'If (Cells(filaDobleClicktz4, columnaDobleClicktz4 + 1) = "") Then
'dato_validacion.Text = "Ingresar fuente de información"
'dato_validacion.BackColor = RGB(255, 255, 0)
'End If


Application.EnableEvents = True

End Sub












