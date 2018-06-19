VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_tz2 
   Caption         =   "Formulario de relevamiento - Trazadora 2"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13410
   OleObjectBlob   =   "userForm_tz2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_tz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_tz2_verificacion_blancos = 1) Then
    MsgBox ("No se han completado todos los campos")
End If


If (auxiliarInexistenteTz2 = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

    auxiliarInexistenteTz2 = 0

End If


Call userForm_tz2_guardar_datos(filaDobleClickTz2)

MsgBox ("Se han guardado con exito")


If (userForm_tz2.dato_validacion.Text = "Labrar acta" Or userForm_tz2.dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones") Then

    Cells(filaDobleClickTz2, columnaDobleClickTz2).Value = "Labrar acta"

ElseIf (userForm_tz2_verificacion_blancos <> 1) Then

    Cells(filaDobleClickTz2, columnaDobleClickTz2).Value = "Completo"

Else

    Cells(filaDobleClickTz2, columnaDobleClickTz2).Value = "Incompleto"

End If

Unload Me

End Sub

Private Sub dato_control_1_completo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_control_1_completo.Text <> "Si" And dato_control_1_completo.Text <> "SI" And dato_control_1_completo.Text <> "No" And dato_control_1_completo.Text <> "NO" And _
dato_control_1_completo.Text <> "No consta control" And dato_control_1_completo.Text <> "Dato no obligatorio") Then

    dato_control_1_completo.Text = ""

End If

End Sub

Private Sub dato_control_2_completo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_control_2_completo.Text <> "Si" And dato_control_2_completo.Text <> "SI" And dato_control_2_completo.Text <> "No" And dato_control_2_completo.Text <> "NO" And _
dato_control_2_completo.Text <> "No consta control" And dato_control_2_completo.Text <> "Dato no obligatorio") Then

    dato_control_2_completo.Text = ""

End If

End Sub

Private Sub dato_control_3_completo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_control_3_completo.Text <> "Si" And dato_control_3_completo.Text <> "SI" And dato_control_3_completo.Text <> "No" And dato_control_3_completo.Text <> "NO" And _
dato_control_3_completo.Text <> "No consta control" And dato_control_3_completo.Text <> "Dato no obligatorio") Then

    dato_control_3_completo.Text = ""

End If

End Sub

Private Sub dato_control_4_completo_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_control_4_completo.Text <> "Si" And dato_control_4_completo.Text <> "SI" And dato_control_4_completo.Text <> "No" And dato_control_4_completo.Text <> "NO" And _
dato_control_4_completo.Text <> "No consta control" And dato_control_4_completo.Text <> "Dato no obligatorio") Then

    dato_control_4_completo.Text = ""

End If

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_tz2.dato_observaciones.MultiLine = True

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
auxiliarInexistenteTz2 = 0

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fuente.Text = "HCPB" Or dato_fuente.Text = "SIP" Or dato_fuente.Text = "HCA" Or dato_fuente.Text = "PP" Or _
dato_fuente.Text = "No consta fuente de información" Or dato_fuente.Text = "Prestación inexistente") Then


    'sentencias para que se complete el textbox de verificacion de casos as modificar este textbox
    If dato_fuente.Text <> "No consta fuente de información" And dato_fuente.Text <> "Prestación inexistente" And dato_fuente.Text <> "Caso duplicado" Then

        dato_validacion.Text = "Ok"
        dato_validacion.BackColor = RGB(87, 166, 57)
        Call userForm_tz2_permitir_campos_requeridos
        auxiliarInexistenteTz2 = 0
        
    Else
        'con estos 3 if le doy valor y formato al textbox de verificacion de casos
        'ademas se bloquean todas los campos y se pone "Dato no obligatorio"
        If (dato_fuente.Text = "No consta fuente de información") Then
            dato_validacion.Text = "Labrar acta"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz2_dato_no_obligatorio

        ElseIf (dato_fuente.Text = "Prestación inexistente") Then
            dato_validacion.Text = "Labrar acta e indicar fuente de información en observaciones"
            dato_validacion.BackColor = RGB(255, 0, 0)
            Call userform_tz2_dato_no_obligatorio
            auxiliarInexistenteTz2 = 1


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
.AddItem "HCPB"
.AddItem "SIP"
.AddItem "HCA"
.AddItem "PP"
.AddItem "No consta fuente de información"
.AddItem "Prestación inexistente"
End With

'si el control 1 esta completo
With dato_control_1_completo
.AddItem "Si"
.AddItem "No"
.AddItem "No consta control"
End With

'si el control 2 esta completo
With dato_control_2_completo
.AddItem "Si"
.AddItem "No"
.AddItem "No consta control"
End With

'si el control 3 esta completo
With dato_control_3_completo
.AddItem "Si"
.AddItem "No"
.AddItem "No consta control"
End With

'si el control 4 esta completo
With dato_control_4_completo
.AddItem "Si"
.AddItem "No"
.AddItem "No consta control"
End With


''Verifica si hay un valor en la celda de fuente de informacion y si es falso pone la leyenda
''"Ingresar fuente de informacion"
'If (Cells(filaDobleClickTZ2, columnaDobleClickTZ2 + 1) = "") Then
'dato_validacion.Text = "Ingresar fuente de información"
'dato_validacion.BackColor = RGB(255, 255, 0)
'End If


Application.EnableEvents = True

End Sub












