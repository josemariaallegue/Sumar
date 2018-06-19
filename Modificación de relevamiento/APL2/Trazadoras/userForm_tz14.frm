VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userForm_tz14 
   Caption         =   "Formulario de relevamiento - Trazadora 14"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8880
   OleObjectBlob   =   "userForm_tz14.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "userForm_tz14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton4_Click()

Dim fuenteInformacion As String

'verifica si hay blancos y si es verdadero ejecuta un mensaje
If (userForm_tz14_verificacion_blancos = 1) Then
    MsgBox ("No se han completado todos los campos")
End If


If (auxiliarInexistenteTz14 = 1) Then

    fuenteInformacion = InputBox("Por favor ingrese la fuente de información. Seleccione 'cancelar' si ya lo ha hecho con anterioridad", "Fuente de información")

    If (dato_observaciones.Text <> "") Then
    
        dato_observaciones.Text = dato_observaciones.Text & ". " & fuenteInformacion
    
    Else
        
        dato_observaciones.Text = fuenteInformacion
    
End If

    auxiliarInexistenteTz14 = 0

End If


Call userForm_tz14_guardar_datos(filaDobleClickTz14)

MsgBox ("Se han guardado con exito")


If (userForm_tz14_verificacion_blancos <> 1) Then

    Cells(filaDobleClickTz14, columnaDobleClickTz14).Value = "Completo"

Else

    ActiveSheet.Cells(filaDobleClickTz14, columnaDobleClickTz14).Value = "Incompleto"

End If

Unload Me

End Sub

Private Sub dato_diagnostico_Change()

dato_diagnostico.MultiLine = True

End Sub

Private Sub dato_fecha_comite_pregunta_Change()

'con este if se evita que se ingresen datos que no estan permitidos
If (dato_fecha_comite_pregunta.Text <> "Dato no obligatorio" And dato_fecha_comite_pregunta.Text <> "Si" And dato_fecha_comite_pregunta.Text <> "si" And dato_fecha_comite_pregunta.Text <> "No" _
And dato_fecha_comite_pregunta.Text <> "no") Then

    dato_fecha_comite_pregunta.Text = ""

End If

'bloquea y desbloquea el campos de diagnostico encontrado en terreno dependiendo del valor ingresado
If (dato_fecha_comite_pregunta.Text = "Si" Or dato_fecha_comite_pregunta.Text = "si") Then

    With dato_fecha_comite_terreno
        .Text = "Dato no obligatorio"
        .BackColor = RGB(169, 169, 169)
        .Locked = True
    End With

ElseIf (dato_fecha_comite_pregunta.Text = "No" Or dato_fecha_comite_pregunta.Text = "no") Then
    
    With dato_fecha_comite_terreno
        .Locked = False
        If (.Text = "Dato no obligatorio") Then
            .Text = ""
        End If
        .BackColor = RGB(255, 255, 255)
    End With

End If
End Sub

Private Sub dato_fecha_comite_terreno_Change()

End Sub

Private Sub dato_observaciones_Change()

'para que siga escribiendo en una linea de inferior
userForm_tz14.dato_observaciones.MultiLine = True

auxiliarGuardadotz14tz14 = 1

End Sub


Private Sub dato_validacion_Change()

'para que siga escribiendo en una linea de inferior
dato_validacion.MultiLine = True

auxiliarGuardadotz14 = 1

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

'pregunta respecto a la entrada de comite encontrada
With dato_fecha_comite_pregunta
.AddItem "Si"
.AddItem "No"
End With


Application.EnableEvents = True

End Sub

