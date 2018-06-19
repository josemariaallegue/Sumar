Attribute VB_Name = "generales"

'brief oculta y protege celdas determinadas en las solapas de relevamiento
'param void
'return void

Sub ocultar_y_proteger()


Dim contraseña As String

contraseña = "crowe2017"

ThisWorkbook.Sheets("Niños y Adolescentes").Range("N:N,Q:AE,AG:AG,AI:AT").EntireColumn.Hidden = True
ThisWorkbook.Sheets("Niños y Adolescentes").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True
    
ThisWorkbook.Sheets("Adultos").Range("N:N,Q:AE,AG:AG,AI:AU").EntireColumn.Hidden = True
ThisWorkbook.Sheets("Adultos").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True
    
ThisWorkbook.Sheets("Embarazos y Partos").Range("N:N,Q:AK,AM:AM,AO:AZ").EntireColumn.Hidden = True
ThisWorkbook.Sheets("Embarazos y Partos").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True
    
ThisWorkbook.Sheets("Niños en internación").Range("N:N,Q:V,X:X,Z:AK").EntireColumn.Hidden = True
ThisWorkbook.Sheets("Niños en internación").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True
    
ThisWorkbook.Sheets("Embarazos de alto riesgo").Range("N:N,Q:AA,AC:AC,AE:Ap").EntireColumn.Hidden = True
ThisWorkbook.Sheets("Embarazos de alto riesgo").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

End Sub

'brief muestra y desprotege celdas determinadas en las solapas de relevamiento
'param void
'return void

Sub mostrar_y_desproteger()

Dim contraseña As String

contraseña = "crowe2017"

If ((InputBox("Ingrese la contraseña", "Desprotección") = contraseña)) Then

    ThisWorkbook.Sheets("Niños y Adolescentes").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("Niños y Adolescentes").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("Adultos").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("Adultos").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("Embarazos y Partos").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("Embarazos y Partos").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("Niños en internación").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("Niños en internación").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("Embarazos de alto riesgo").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("Embarazos de alto riesgo").Range("A:AZ").EntireColumn.Hidden = False
    
Else

MsgBox "Se ha ingresado una contraseña erronea"

End If

End Sub


'brief: Analisa el formulario (viendo solo los motivos 1, 2 y 3). El motivo 4 debe ser verificado a mano
'param: void
'return: void


Function analisis_niños_adolescentes()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'11 es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ThisWorkbook.Sheets("Niños y Adolescentes").Cells(6, 42).Value = Application.UserName
ThisWorkbook.Sheets("Niños y Adolescentes").Cells(6, 43).Value = Date


'hace las siguientes lineas hasta que encuentra en la columna 13 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 13).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias
    ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 41).Value = ""
    ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 43).Value = ""
    
    'los primeros 3 if son para los motivos 1, 2 y 3 respectivamente
    If (ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 15).Value = "A") Then
    
        ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 41).Value = 1
        ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 43).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 15).Value = "B") Then
    
        ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 41).Value = 2
        ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 43).Value = "Prestación inexistente"
    
    ElseIf (ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 15).Value = "C") Then
    
        ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 41).Value = 5
        ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 43).Value = "Fuente invalida"
      
    'para el motivo 4
    ElseIf (ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 13).Value = "Incompleto" Or ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 13).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
    
        'recorre la fila viendo que celdas estan vacias o dicen no y completa
        For j = 17 To 31
        
            If ((ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, j).Value = "No") And _
            ThisWorkbook.Sheets("Niños y Adolescentes").Cells(10, j).Value <> "Diagnóstico") Then
                
                If (flag = 0) Then
                
                    leyenda = leyenda & ThisWorkbook.Sheets("Niños y Adolescentes").Cells(10, j).Value
                    ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 41).Value = 3
                    flag = 1
                    
                Else
                
                    leyenda = leyenda & ", " & ThisWorkbook.Sheets("Niños y Adolescentes").Cells(10, j).Value
                    
                End If
            
            End If
        
        Next
        
        
        If (flag = 1) Then
        
            ThisWorkbook.Sheets("Niños y Adolescentes").Cells(i, 43).Value = leyenda
            
        End If
        
    Else
    
        MsgBox ("Niños y adolescentes. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Function

'brief: Analisa el formulario (viendo solo los motivos 1, 2 y 3). El motivo 4 debe ser verificado a mano
'param: void
'return: void


Function analisis_adultos()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'11 es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ThisWorkbook.Sheets("Adultos").Cells(6, 42).Value = Application.UserName
ThisWorkbook.Sheets("Adultos").Cells(6, 43).Value = Date


'hace las siguientes lineas hasta que encuentra en la columna 13 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("Adultos").Cells(i, 13).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias
    ThisWorkbook.Sheets("Adultos").Cells(i, 41).Value = ""
    ThisWorkbook.Sheets("Adultos").Cells(i, 43).Value = ""
    
    'los primeros 3 if son para los motivos 1, 2 y 3 respectivamente
    If (ThisWorkbook.Sheets("Adultos").Cells(i, 15).Value = "A") Then
    
        ThisWorkbook.Sheets("Adultos").Cells(i, 41).Value = 1
        ThisWorkbook.Sheets("Adultos").Cells(i, 43).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("Adultos").Cells(i, 15).Value = "B") Then
    
        ThisWorkbook.Sheets("Adultos").Cells(i, 41).Value = 2
        ThisWorkbook.Sheets("Adultos").Cells(i, 43).Value = "Prestación inexistente"
        
    ElseIf (ThisWorkbook.Sheets("Adultos").Cells(i, 15).Value = "C") Then
    
        ThisWorkbook.Sheets("Adultos").Cells(i, 41).Value = 5
        ThisWorkbook.Sheets("Adultos").Cells(i, 43).Value = "Fuente invalida"
      
    'para el motivo 4
    ElseIf (ThisWorkbook.Sheets("Adultos").Cells(i, 13).Value = "Incompleto" Or ThisWorkbook.Sheets("Adultos").Cells(i, 13).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
    
        'recorre la fila viendo que celdas estan vacias o dicen no y completa
        For j = 17 To 31
        
            If ((ThisWorkbook.Sheets("Adultos").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("Adultos").Cells(i, j).Value = "No") And _
            ThisWorkbook.Sheets("Adultos").Cells(10, j).Value <> "Diagnóstico") Then
                
                If (flag = 0) Then
                
                    leyenda = leyenda & ThisWorkbook.Sheets("Adultos").Cells(10, j).Value
                    ThisWorkbook.Sheets("Adultos").Cells(i, 41).Value = 3
                    flag = 1
                    
                Else
                
                    leyenda = leyenda & ", " & ThisWorkbook.Sheets("Adultos").Cells(10, j).Value
                    
                End If
            
            End If
        
        Next
        
        
        If (flag = 1) Then
        
            ThisWorkbook.Sheets("Adultos").Cells(i, 43).Value = leyenda
        End If
        
    Else
    
        MsgBox ("Adultos. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop
End Function

'brief: Analisa el formulario (viendo solo los motivos 1, 2 y 3). El motivo 4 debe ser verificado a mano
'param: void
'return: void


Function analisis_embarazos_partos()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'11 es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ThisWorkbook.Sheets("Embarazos y Partos").Cells(6, 48).Value = Application.UserName
ThisWorkbook.Sheets("Embarazos y Partos").Cells(6, 49).Value = Date


'hace las siguientes lineas hasta que encuentra en la columna 13 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 13).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias
    ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 47).Value = ""
    ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 49).Value = ""
    
    'los primeros 3 if son para los motivos 1, 2 y 3 respectivamente
    If (ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 15).Value = "A") Then
    
        ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 47).Value = 1
        ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 49).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 15).Value = "B") Then
    
        ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 47).Value = 2
        ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 49).Value = "Prestación inexistente"
    
    ElseIf (ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 15).Value = "C") Then
    
        ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 47).Value = 5
        ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 49).Value = "Fuente invalida"
      
    'para el motivo 4
    ElseIf (ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 13).Value = "Incompleto" Or ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 13).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
    
        'recorre la fila viendo que celdas estan vacias o dicen no y completa
        For j = 18 To 37
        
            If ((ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, j).Value = "No") And _
            ThisWorkbook.Sheets("Embarazos y Partos").Cells(10, j).Value <> "Diagnóstico") Then
                
                If (flag = 0) Then
                
                    leyenda = leyenda & ThisWorkbook.Sheets("Embarazos y Partos").Cells(10, j).Value
                    ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 47).Value = 3
                    flag = 1
                    
                Else
                
                    leyenda = leyenda & ", " & ThisWorkbook.Sheets("Embarazos y Partos").Cells(10, j).Value
                    
                End If
            
            End If
        
        Next
        
        
        If (flag = 1) Then
        
            ThisWorkbook.Sheets("Embarazos y Partos").Cells(i, 49).Value = leyenda
            
        End If
        
    Else
    
        MsgBox ("Embarazos y partos. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Function

'brief: Analisa el formulario (viendo solo los motivos 1, 2 y 3). El motivo 4 debe ser verificado a mano
'param: void
'return: void


Function analisis_niños_internacion()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'11 es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ThisWorkbook.Sheets("Niños en internación").Cells(6, 33).Value = Application.UserName
ThisWorkbook.Sheets("Niños en internación").Cells(6, 34).Value = Date


'hace las siguientes lineas hasta que encuentra en la columna 13 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("Niños en internación").Cells(i, 13).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias
    ThisWorkbook.Sheets("Niños en internación").Cells(i, 32).Value = ""
    ThisWorkbook.Sheets("Niños en internación").Cells(i, 34).Value = ""
    
    'los primeros 3 if son para los motivos 1, 2 y 3 respectivamente
    If (ThisWorkbook.Sheets("Niños en internación").Cells(i, 15).Value = "A") Then
    
        ThisWorkbook.Sheets("Niños en internación").Cells(i, 32).Value = 1
        ThisWorkbook.Sheets("Niños en internación").Cells(i, 34).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("Niños en internación").Cells(i, 15).Value = "B") Then
    
        ThisWorkbook.Sheets("Niños en internación").Cells(i, 32).Value = 2
        ThisWorkbook.Sheets("Niños en internación").Cells(i, 34).Value = "Prestación inexistente"
    
    ElseIf (ThisWorkbook.Sheets("Niños en internación").Cells(i, 15).Value = "C") Then

        ThisWorkbook.Sheets("Niños en internación").Cells(i, 32).Value = 5
        ThisWorkbook.Sheets("Niños en internación").Cells(i, 34).Value = "Fuente invalida"
      
    'para el motivo 4
    ElseIf (ThisWorkbook.Sheets("Niños en internación").Cells(i, 13).Value = "Incompleto" Or ThisWorkbook.Sheets("Niños en internación").Cells(i, 13).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0

        'recorre la fila viendo que celdas estan vacias o dicen no y completa
        For j = 17 To 22
        
            If ((ThisWorkbook.Sheets("Niños en internación").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("Niños en internación").Cells(i, j).Value = "No") And _
            ThisWorkbook.Sheets("Niños en internación").Cells(10, j).Value <> "Diagnóstico") Then
                
                If (flag = 0) Then
                
                    leyenda = leyenda & ThisWorkbook.Sheets("Niños en internación").Cells(10, j).Value
                    ThisWorkbook.Sheets("Niños en internación").Cells(i, 32).Value = 3
                    flag = 1
                    
                Else
                
                    leyenda = leyenda & ", " & ThisWorkbook.Sheets("Niños en internación").Cells(10, j).Value
                    
                End If
            
            End If
        
        Next
        
        
        If (flag = 1) Then
        
            ThisWorkbook.Sheets("Niños en internación").Cells(i, 34).Value = leyenda
            
        End If
        
    Else
    
        MsgBox ("Niños en internarcion. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Function

'brief: Analisa el formulario (viendo solo los motivos 1, 2 y 3). El motivo 4 debe ser verificado a mano
'param: void
'return: void


Function analisis_embarazos_alto_riesgo()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'11 es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(6, 38).Value = Application.UserName
ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(6, 39).Value = Date

        
'hace las siguientes lineas hasta que encuentra en la columna 13 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 13).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias
    ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 37).Value = ""
    ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 39).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 15).Value = "A") Then
    
        ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 37).Value = 1
        ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 39).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 15).Value = "B") Then
    
        ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 37).Value = 2
        ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 39).Value = "Prestación inexistente"
        
    ElseIf (ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 15).Value = "C") Then
    
        ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 37).Value = 5
        ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 39).Value = "Prestación inexistente"
      
    'para el motivo 4
    ElseIf (ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 13).Value = "Incompleto" Or ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 13).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
        
        'recorre la fila viendo que celdas estan vacias o dicen no y completa
        For j = 17 To 27
        
            If ((ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, j).Value = "No") And _
            ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(10, j).Value <> "Diagnóstico") Then
                
                If (flag = 0) Then
                
                    leyenda = leyenda & ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(10, j).Value
                    ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 37).Value = 3
                    flag = 1
                    
                Else
                
                    leyenda = leyenda & ", " & ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(10, j).Value
                    
                End If
            
            End If
        
        Next
        
        
        If (flag = 1) Then
        
            ThisWorkbook.Sheets("Embarazos de alto riesgo").Cells(i, 39).Value = leyenda
            
        End If
        
    Else
    
        MsgBox ("Embarazos de alto riesgo. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Function


Sub analisis()

On Error Resume Next

    Call analisis_niños_adolescentes
    Call analisis_adultos
    Call analisis_embarazos_partos
    Call analisis_niños_internacion
    Call analisis_embarazos_alto_riesgo
    
On Error GoTo 0

End Sub

