Attribute VB_Name = "Generales"

'brief oculta y protege celdas determinadas en las solapas de relevamiento
'param void
'return void

Sub ocultar_y_proteger()


Dim contraseña As String

'determina el valor de la contraseña
contraseña = "crowe2017"

ThisWorkbook.Sheets("TZ1").Range("K:K,M:W").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ1").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True
    
ThisWorkbook.Sheets("TZ2").Range("J:J,L:L,N:N,P:P,R:R,T:AH").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ2").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ3").Range("J:J,L:Y").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ3").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ4").Range("K:K,M:N,P:Y").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ4").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ7").Range("J:J,M:X").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ7").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ8").Range("J:J,M:N,P:AA").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ8").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ9").Range("J:J,M:N,P:Q,S:AE").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ9").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ10").Range("K:K,M:V").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ10").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ11").Range("K:K,O:V").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ11").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ12").Range("J:J,M:O,Q:AB").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ12").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ13").Range("J:J,M:N,P:AE").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ13").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

ThisWorkbook.Sheets("TZ14").Range("H:Q").EntireColumn.Hidden = True
ThisWorkbook.Sheets("TZ14").Protect Password:=contraseña, _
DrawingObjects:=False, Contents:=True, Scenarios:=False, _
AllowFormattingCells:=True, AllowFormattingColumns:=False, _
AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
AllowUsingPivotTables:=True, UserInterfaceOnly:=True

End Sub

'brief muestra y desprotege celdas determinadas en las solapas de relevamiento
'param void
'return void

Sub mostrar_y_desproteger()

Dim contraseña As String

'determina el valor de la contraseña
contraseña = "crowe2017"

If ((InputBox("Ingrese la contraseña", "Desprotección") = contraseña)) Then

    ThisWorkbook.Sheets("TZ1").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ1").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ2").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ2").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ3").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ3").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ4").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ4").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ7").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ7").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ8").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ8").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ9").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ9").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ10").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ10").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ11").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ11").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ12").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ12").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ13").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ13").Range("A:AZ").EntireColumn.Hidden = False
    
    ThisWorkbook.Sheets("TZ14").Unprotect Password:=contraseña
    ThisWorkbook.Sheets("TZ14").Range("A:AZ").EntireColumn.Hidden = False
    
Else

    MsgBox "Se ha ingresado una contraseña erronea"

End If

End Sub


'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void

Sub analisis_tz1()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("TZ1").Cells(5, 18).Value = Application.UserName
ThisWorkbook.Sheets("TZ1").Cells(5, 19).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("TZ1").Cells(i, 10).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("TZ1").Cells(i, 20).Value = ""
    ThisWorkbook.Sheets("TZ1").Cells(i, 22).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("TZ1").Cells(i, 12).Value = "A") Then
    
        ThisWorkbook.Sheets("TZ1").Cells(i, 20).Value = 1
        ThisWorkbook.Sheets("TZ1").Cells(i, 22).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("TZ1").Cells(i, 12).Value = "B") Then
    
        ThisWorkbook.Sheets("TZ1").Cells(i, 20).Value = 2
        ThisWorkbook.Sheets("TZ1").Cells(i, 22).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("TZ1").Cells(i, 10).Value = "Incompleto" Or ThisWorkbook.Sheets("TZ1").Cells(i, 10).Value = "Completo") Then
        
        'verifico que tanto la edad gestacional por fum como la edad gestacional al control declardo esten fuera de los rangos permitidos
        If (ThisWorkbook.Sheets("TZ1").Cells(i, 18).Value >= 14 And ThisWorkbook.Sheets("TZ1").Cells(i, 14).Value >= 14) Then
            
            ThisWorkbook.Sheets("TZ1").Cells(i, 20).Value = 3
            ThisWorkbook.Sheets("TZ1").Cells(i, 22).Value = "Control prenatal mayor a 13 semanas de gestación"
                
        End If
        
    Else
    
        MsgBox ("TZ1. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub

'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz2()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz2").Cells(5, 28).Value = Application.UserName
ThisWorkbook.Sheets("tz2").Cells(5, 29).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("tz2").Cells(i, 9).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = ""
    ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("tz2").Cells(i, 11).Value = "A") Then
    
        ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = 1
        ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("tz2").Cells(i, 11).Value = "B") Then
    
        ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = 2
        ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("tz2").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("tz2").Cells(i, 9).Value = "Completo") Then
    
        'verifico que solamente el 4 control este incompleto
        If (ThisWorkbook.Sheets("tz2").Cells(i, 20).Value = "No consta control" Or ThisWorkbook.Sheets("tz2").Cells(i, 20).Value = "") Then
        
            ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = 3
            ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = "No se verifican controles completos"
        
        'verifico que no exista control antes de la semana 20
        ElseIf (ThisWorkbook.Sheets("tz2").Cells(i, 25).Value >= 21 And ThisWorkbook.Sheets("tz2").Cells(i, 27).Value >= 21) Then
            
            'se tiene que verificar que la celda no contenga estos valores porque si los llega a tener tira al beneficiaro por este
            'motivo cuando no es correcto
            If (ThisWorkbook.Sheets("tz2").Cells(i, 12).Value <> "" Or ThisWorkbook.Sheets("tz2").Cells(i, 12).Value <> "Dato no obligatorio") Then
            
                ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = 3
                ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = "No se encontró control antes de la semana 20 de gestación"
                
            End If
            
        'verifico que no se encuentre un control despues de la seman 34
        ElseIf (ThisWorkbook.Sheets("tz2").Cells(i, 26).Value < 34 And ThisWorkbook.Sheets("tz2").Cells(i, 21).Value < 34) Then
        
            If (ThisWorkbook.Sheets("tz2").Cells(i, 12).Value <> "" Or ThisWorkbook.Sheets("tz2").Cells(i, 12).Value <> "Dato no obligatorio") Then
            
                'se tiene que verificar que la celda no contenga estos valores porque si los llega a tener tira al beneficiaro por este
                'motivo cuando no es correcto
                ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = 3
                ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = "No se encontró control después de la semana 34 de gestación"
            
            End If
        
        'verifico que no se cumpla la diferencia de 8 dias entre alguno de los controles
        ElseIf (ThisWorkbook.Sheets("tz2").Cells(i, 28).Value < 8 Or ThisWorkbook.Sheets("tz2").Cells(i, 29).Value < 8 Or _
        ThisWorkbook.Sheets("tz2").Cells(i, 30).Value < 8) Then
            
            ThisWorkbook.Sheets("tz2").Cells(i, 31).Value = 3
            ThisWorkbook.Sheets("tz2").Cells(i, 33).Value = "No cumple con la diferencia de 8 días entre los controles"
            
        End If

        
    Else
    
        MsgBox ("tz2. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub

'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz3()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz3").Cells(5, 21).Value = Application.UserName
ThisWorkbook.Sheets("tz3").Cells(5, 22).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("tz3").Cells(i, 9).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = ""
    ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("tz3").Cells(i, 11).Value = "A") Then
    
        ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = 1
        ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("tz3").Cells(i, 11).Value = "B") Then
    
        ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = 2
        ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("tz3").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("tz3").Cells(i, 9).Value = "Completo") Then
    
        'verifico que la que la celda de peso sea menor a al rango permitido y que no este vacia
        If (ThisWorkbook.Sheets("tz3").Cells(i, 12).Value < 750 And ThisWorkbook.Sheets("tz3").Cells(i, 12) <> "") Then
        
            ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = 3
            ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = "El peso del niño al nacer es menor a 750 grs."
        
        'verifico que la que la celda de peso sea mayor a al rango permitido y que no este vacia
        ElseIf (ThisWorkbook.Sheets("tz3").Cells(i, 12).Value >= 1501 And ThisWorkbook.Sheets("tz3").Cells(i, 12) <> "") Then
            
            ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = 3
            ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = "El peso del niño al nacer es mayor a 1.500 Grs"
        
        'si la celda de peso esta vacia lo tiro por dato incompleto
        ElseIf (ThisWorkbook.Sheets("tz3").Cells(i, 12) = "") Then
                    
            ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = 3
            ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = "Datos incompletos: peso"
        
        'verifico que el paciente no haya fallecido
        ElseIf (ThisWorkbook.Sheets("tz3").Cells(i, 14).Value = "Fallecido") Then
            
            ThisWorkbook.Sheets("tz3").Cells(i, 22).Value = 3
            ThisWorkbook.Sheets("tz3").Cells(i, 24).Value = "No cumple con la sobrevida de 28 días"
            
        End If
    Else
    
        MsgBox ("tz3. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub

'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz4()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 11

'marca quien y cuando se analizo
ThisWorkbook.Sheets("TZ4").Cells(6, 22).Value = Application.UserName
ThisWorkbook.Sheets("TZ4").Cells(6, 23).Value = Date

'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("TZ4").Cells(i, 10).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("TZ4").Cells(i, 22).Value = ""
    ThisWorkbook.Sheets("TZ4").Cells(i, 24).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("TZ4").Cells(i, 12).Value = "A") Then
    
        ThisWorkbook.Sheets("TZ4").Cells(i, 22).Value = 1
        ThisWorkbook.Sheets("TZ4").Cells(i, 24).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("TZ4").Cells(i, 12).Value = "B") Then
    
        ThisWorkbook.Sheets("TZ4").Cells(i, 22).Value = 2
        ThisWorkbook.Sheets("TZ4").Cells(i, 24).Value = "Prestación inexistente"
    
    'para el motivo 3

    ElseIf (ThisWorkbook.Sheets("TZ4").Cells(i, 10).Value = "Incompleto" Or ThisWorkbook.Sheets("TZ4").Cells(i, 10).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
        
        'verifica que no haya errores, como que no se haya modificado la fecha de finalizacion del periodo
        If (ThisWorkbook.Sheets("tz4").Cells(i, 21).Value <> "Modificar fecha") Then
            
            'los primeros 2 verifica que la antiguedad este entre los rangos permitidos
            If (ThisWorkbook.Sheets("tz4").Cells(i, 21).Value > 10) Then
            
                ThisWorkbook.Sheets("tz4").Cells(i, 22).Value = 3
                ThisWorkbook.Sheets("tz4").Cells(i, 24).Value = "El niño es mayor a 10 años"
            
            ElseIf (ThisWorkbook.Sheets("tz4").Cells(i, 21).Value < 0) Then
            
                ThisWorkbook.Sheets("tz4").Cells(i, 22).Value = "3"
                ThisWorkbook.Sheets("tz4").Cells(i, 24).Value = "La edad del niño es menor a 0"
            
            'si esta dentro del rango verifica los datos obligatorios
            Else
            
                For j = 13 To 17

                    If ((ThisWorkbook.Sheets("TZ4").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("TZ4").Cells(i, j).Value = "No") And _
                    ThisWorkbook.Sheets("TZ4").Cells(10, j).Value <> "¿Corresponde completar: Perímetro Cefálico?") Then

                        If (flag = 0) Then

                            leyenda = leyenda & ThisWorkbook.Sheets("TZ4").Cells(10, j).Value
                            flag = 1

                         Else

                            leyenda = leyenda & ", " & ThisWorkbook.Sheets("TZ4").Cells(10, j).Value

                        End If

                    End If

                Next


                If (flag = 1) Then
                    
                    ThisWorkbook.Sheets("TZ4").Cells(i, 22).Value = 3
                    ThisWorkbook.Sheets("TZ4").Cells(i, 24).Value = leyenda

                End If
            
                
            End If
            
        End If
        
    Else
    
        MsgBox ("TZ4. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub


'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz7()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String

'es la fila donde el auditor comienza a relevar
i = 12

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz7").Cells(5, 22).Value = Application.UserName
ThisWorkbook.Sheets("tz7").Cells(5, 23).Value = Date

'hace las siguientes lineas hasta que encuentra en la columna 9 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("tz7").Cells(i, 9).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz7").Cells(i, 21).Value = ""
    ThisWorkbook.Sheets("tz7").Cells(i, 23).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("tz7").Cells(i, 11).Value = "A") Then
    
        ThisWorkbook.Sheets("tz7").Cells(i, 21).Value = 1
        ThisWorkbook.Sheets("tz7").Cells(i, 23).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("tz7").Cells(i, 11).Value = "B") Then
    
        ThisWorkbook.Sheets("tz7").Cells(i, 21).Value = 2
        ThisWorkbook.Sheets("tz7").Cells(i, 23).Value = "Prestación inexistente"
    
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("tz7").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("tz7").Cells(i, 9).Value = "Completo") Then
        
        'verifica que no haya errores, como que no se haya modificado la fecha o que falte alguna fecha
        If (ThisWorkbook.Sheets("tz7").Cells(i, 19).Value <> "Modificar fecha" And ThisWorkbook.Sheets("tz7").Cells(i, 19).Value <> "Se coloco que la fechas no coinciden pero no completo con la fecha") Then
            
            'los primeros 2 verifica que la antiguedad este entre los rangos permitidos
            If (ThisWorkbook.Sheets("tz7").Cells(i, 19).Value > 1) Then
            
                ThisWorkbook.Sheets("tz7").Cells(i, 21).Value = 3
                ThisWorkbook.Sheets("tz7").Cells(i, 23).Value = "La determinacion de TSOMF supera el año"
            
            ElseIf (ThisWorkbook.Sheets("tz7").Cells(i, 19).Value < 0) Then
            
                ThisWorkbook.Sheets("tz7").Cells(i, 21).Value = "3"
                ThisWorkbook.Sheets("tz7").Cells(i, 23).Value = "La antiguedad de la realizacion de TSOMF es menor a 0"
            
            'si esta dentro del rango verifica los datos obligatorios
            Else
            
                'depura los valores cada vez que entra
                leyenda = "Datos incompletos: "
                
                If (ThisWorkbook.Sheets("tz7").Cells(i, 15).Value = "") Then
                
                    leyenda = leyenda & ThisWorkbook.Sheets("tz7").Cells(11, 15).Value
                
                    ThisWorkbook.Sheets("tz7").Cells(i, 21).Value = 3
                    ThisWorkbook.Sheets("tz7").Cells(i, 23).Value = leyenda
                    
                End If
                
            End If
            
        End If
    Else
    
        MsgBox ("tz7. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub

'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz8()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("TZ8").Cells(5, 24).Value = Application.UserName
ThisWorkbook.Sheets("TZ8").Cells(5, 25).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("TZ8").Cells(i, 9).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("TZ8").Cells(i, 24).Value = ""
    ThisWorkbook.Sheets("TZ8").Cells(i, 26).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("TZ8").Cells(i, 11).Value = "A") Then
    
        ThisWorkbook.Sheets("TZ8").Cells(i, 24).Value = 1
        ThisWorkbook.Sheets("TZ8").Cells(i, 26).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("TZ8").Cells(i, 11).Value = "B") Then
    
        ThisWorkbook.Sheets("TZ8").Cells(i, 24).Value = 2
        ThisWorkbook.Sheets("TZ8").Cells(i, 26).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("TZ8").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("TZ8").Cells(i, 9).Value = "Completo") Then
    
        'verifico que las vacunas esten dadas entre los rangos posibles (todas las combinaciones posibles)
        If ((ThisWorkbook.Sheets("TZ8").Cells(i, 22).Value > 24 Or ThisWorkbook.Sheets("TZ8").Cells(i, 22).Value < 15) And _
        (ThisWorkbook.Sheets("TZ8").Cells(i, 23).Value > 24 Or ThisWorkbook.Sheets("TZ8").Cells(i, 23).Value < 15)) Then
        
            ThisWorkbook.Sheets("TZ8").Cells(i, 24).Value = 3
            ThisWorkbook.Sheets("TZ8").Cells(i, 26).Value = "Ninguna de las vacunas fueron aplicadas entre los 15 y 24 meses"
            
        ElseIf (ThisWorkbook.Sheets("TZ8").Cells(i, 22).Value > 24 Or ThisWorkbook.Sheets("TZ8").Cells(i, 22).Value < 15) Then
        
            ThisWorkbook.Sheets("TZ8").Cells(i, 24).Value = 3
            ThisWorkbook.Sheets("TZ8").Cells(i, 26).Value = "La vacuna Cuádruple Bacteriana o Pentavalente no fue aplicada entre los 15 y 24 meses"
            
        ElseIf (ThisWorkbook.Sheets("TZ8").Cells(i, 23).Value > 24 Or ThisWorkbook.Sheets("TZ8").Cells(i, 23).Value < 15) Then
            
            ThisWorkbook.Sheets("TZ8").Cells(i, 24).Value = 3
            ThisWorkbook.Sheets("TZ8").Cells(i, 26).Value = "La vacuna Antipoliomielítica no fue aplicada entre los 15 y 24 meses"
        End If
        
    Else
    
        MsgBox ("TZ8. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub

'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz9()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz9").Cells(5, 28).Value = Application.UserName
ThisWorkbook.Sheets("tz9").Cells(5, 29).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("tz9").Cells(i, 9).Value = ""

    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = ""
    ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = ""

    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("tz9").Cells(i, 11).Value = "A") Then

        ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 1
        ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "No consta fuente de información"

    ElseIf (ThisWorkbook.Sheets("tz9").Cells(i, 11).Value = "B") Then

        ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 2
        ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "Prestación inexistente"

    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("tz9").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("tz9").Cells(i, 9).Value = "Completo") Then

        'verifico que las vacunas fueron dadas entre los rangos permitidos (todas las combinaciones posibles)
        If ((ThisWorkbook.Sheets("tz9").Cells(i, 25).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 25).Value < 5) And _
        (ThisWorkbook.Sheets("tz9").Cells(i, 26).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 26).Value < 5) And _
        (ThisWorkbook.Sheets("tz9").Cells(i, 27).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 27).Value < 5)) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "Ninguna de las vacunas fueron aplicadas entre los 5 y 7 años"

        ElseIf ((ThisWorkbook.Sheets("tz9").Cells(i, 25).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 25).Value < 5) And _
        (ThisWorkbook.Sheets("tz9").Cells(i, 26).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 26).Value < 5)) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "La vacunas Triple Bacteriana y Triple Viral no fueron aplicadas entre los 5 y 7 años"

        ElseIf ((ThisWorkbook.Sheets("tz9").Cells(i, 25).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 25).Value < 5) And _
        (ThisWorkbook.Sheets("tz9").Cells(i, 27).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 27).Value < 5)) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "La vacunas Triple Bacteriana y Antipoliomielítica no fueron aplicadas entre los 5 y 7 años"

        ElseIf ((ThisWorkbook.Sheets("tz9").Cells(i, 26).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 26).Value < 5) And _
        (ThisWorkbook.Sheets("tz9").Cells(i, 27).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 27).Value < 5)) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "La vacunas Triple Viral y Antipoliomielítica no fueron aplicadas entre los 5 y 7 años"

        ElseIf (ThisWorkbook.Sheets("tz9").Cells(i, 25).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 25).Value < 5) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "La vacuna Triple Bacteriana no fue aplicada entre los 5 y 7 años"

        ElseIf (ThisWorkbook.Sheets("tz9").Cells(i, 26).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 26).Value < 5) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "La vacuna Triple Viral o Doble Viral no fue aplicada entre los 5 y 7 años"

        ElseIf (ThisWorkbook.Sheets("tz9").Cells(i, 27).Value > 7 Or ThisWorkbook.Sheets("tz9").Cells(i, 27).Value < 5) Then

            ThisWorkbook.Sheets("tz9").Cells(i, 28).Value = 3
            ThisWorkbook.Sheets("tz9").Cells(i, 30).Value = "La vacuna  Antipoliomielítica (5° refuerzo) no fue aplicada entre los 5 y 7 años"

        End If

    Else

        MsgBox ("tz9. Hubo un error en la fila: " & i & ". Verificar con el che pibe")

    End If

    i = i + 1

Loop

End Sub

'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz10()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz10").Cells(5, 19).Value = Application.UserName
ThisWorkbook.Sheets("tz10").Cells(5, 20).Value = Date

'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("tz10").Cells(i, 10).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz10").Cells(i, 19).Value = ""
    ThisWorkbook.Sheets("tz10").Cells(i, 21).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("tz10").Cells(i, 12).Value = "A") Then
    
        ThisWorkbook.Sheets("tz10").Cells(i, 19).Value = 1
        ThisWorkbook.Sheets("tz10").Cells(i, 21).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("tz10").Cells(i, 12).Value = "B") Then
    
        ThisWorkbook.Sheets("tz10").Cells(i, 19).Value = 2
        ThisWorkbook.Sheets("tz10").Cells(i, 21).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("tz10").Cells(i, 10).Value = "Incompleto" Or ThisWorkbook.Sheets("tz10").Cells(i, 10).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
            
        For j = 13 To 15

            If (ThisWorkbook.Sheets("tz10").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("tz10").Cells(i, j).Value = "No") Then

                If (flag = 0) Then

                    leyenda = leyenda & ThisWorkbook.Sheets("tz10").Cells(9, j).Value
                    flag = 1

                 Else

                    leyenda = leyenda & ", " & ThisWorkbook.Sheets("tz10").Cells(9, j).Value

                End If

            End If

        Next


        If (flag = 1) Then
            
            ThisWorkbook.Sheets("tz10").Cells(i, 19).Value = 3
            ThisWorkbook.Sheets("tz10").Cells(i, 21).Value = leyenda

        End If
    Else
    
        MsgBox ("tz10. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub


'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz11()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz11").Cells(5, 20).Value = Application.UserName
ThisWorkbook.Sheets("tz11").Cells(5, 21).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("tz11").Cells(i, 10).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz11").Cells(i, 19).Value = ""
    ThisWorkbook.Sheets("tz11").Cells(i, 21).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("tz11").Cells(i, 12).Value = "A") Then
    
        ThisWorkbook.Sheets("tz11").Cells(i, 19).Value = 1
        ThisWorkbook.Sheets("tz11").Cells(i, 21).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("tz11").Cells(i, 12).Value = "B") Then
    
        ThisWorkbook.Sheets("tz11").Cells(i, 19).Value = 2
        ThisWorkbook.Sheets("tz11").Cells(i, 21).Value = "Prestación inexistente"
    
'    'para el motivo 3
'    ElseIf (ThisWorkbook.Sheets("tz11").Cells(i, 10).Value = "Incompleto" Or ThisWorkbook.Sheets("tz11").Cells(i, 10).Value = "Completo") Then
'
'        'depura los valores cada vez que entra
'        leyenda = "Datos incompletos: "
'
'            'verifica que la celda estan vacias o dicen no y completa
'            If (ThisWorkbook.Sheets("tz11").Cells(i, 15).Value = "" Or ThisWorkbook.Sheets("tz11").Cells(i, 15).Value = "No") Then
'
'                leyenda = leyenda & ThisWorkbook.Sheets("tz11").Cells(9, 15).Value
'                ThisWorkbook.Sheets("tz11").Cells(i, 19).Value = 3
'                ThisWorkbook.Sheets("tz11").Cells(i, 21).Value = leyenda
'
'            End If
        
    Else
    
        MsgBox ("tz11. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub
'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz12()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("TZ12").Cells(5, 26).Value = Application.UserName
ThisWorkbook.Sheets("TZ12").Cells(5, 27).Value = Date

'hace las siguientes lineas hasta que encuentra en la columna 9 (la del doble click) un celda vacia
Do Until ThisWorkbook.Sheets("TZ12").Cells(i, 9).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = ""
    ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("TZ12").Cells(i, 11).Value = "A") Then
    
        ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = 1
        ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("TZ12").Cells(i, 11).Value = "B") Then
    
        ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = 2
        ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("TZ12").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("TZ12").Cells(i, 9).Value = "Completo") Then

        'verifica que no haya errores, como que no se haya modificado la fecha
        If (ThisWorkbook.Sheets("TZ12").Cells(i, 23).Value <> "Modificar fecha") Then
            
            'los primeros 2 verifica que la antiguedad este entre los rangos permitidos
            If (ThisWorkbook.Sheets("TZ12").Cells(i, 23).Value < 25) Then
            
                ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = 3
                ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = "El caso verifica edad menor a los 25 años"
            
            ElseIf (ThisWorkbook.Sheets("TZ12").Cells(i, 23).Value > 64) Then
            
                ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = "3"
                ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = "El caso verifica edad mayor a los 64 años"
            
            'verifica la antiguedad del diagnostico
            ElseIf (ThisWorkbook.Sheets("TZ12").Cells(i, 24).Value > 1) Then
            
                ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = "3"
                ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = "La fecha del diagnóstico supera el año"
            
            'si esta dentro del rango verifica los datos obligatorios
            Else
                        
                    'depura los valores cada vez que entra
                    leyenda = "Datos incompletos: "
                    
                    If (ThisWorkbook.Sheets("TZ12").Cells(i, 15).Value = "" Or ThisWorkbook.Sheets("TZ12").Cells(i, 15).Value = "No consta") Then


                        leyenda = leyenda & "Reporte del resultado histologico"
                        ThisWorkbook.Sheets("TZ12").Cells(i, 25).Value = 3
                        ThisWorkbook.Sheets("TZ12").Cells(i, 27).Value = leyenda
                        
                    End If
            End If
            
        End If
        
    Else
    
        MsgBox ("TZ12. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub


'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz13()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim leyenda As String
Dim flag As Integer

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("TZ13").Cells(5, 29).Value = Application.UserName
ThisWorkbook.Sheets("TZ13").Cells(5, 30).Value = Date

'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("TZ13").Cells(i, 9).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = ""
    ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = ""
    
    'los primeros 2 if son para los motivos 1 y 2 respectivamente
    If (ThisWorkbook.Sheets("TZ13").Cells(i, 11).Value = "A") Then
    
        ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = 1
        ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = "No consta fuente de información"
    
    ElseIf (ThisWorkbook.Sheets("TZ13").Cells(i, 11).Value = "B") Then
    
        ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = 2
        ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = "Prestación inexistente"
    
    'para el motivo 3
    ElseIf (ThisWorkbook.Sheets("TZ13").Cells(i, 9).Value = "Incompleto" Or ThisWorkbook.Sheets("TZ13").Cells(i, 9).Value = "Completo") Then
    
        'depura los valores cada vez que entra
        leyenda = "Datos incompletos: "
        flag = 0
        
        'verifica que no haya errores, como que no se haya modificado la fecha o que falte alguna fecha
        If (ThisWorkbook.Sheets("TZ13").Cells(i, 26).Value <> "Modificar fecha") Then
            
            'los primeros 2 verifica que la antiguedad este entre los rangos permitidos
            If (ThisWorkbook.Sheets("TZ13").Cells(i, 26).Value < 20) Then
            
                ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = 3
                ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = "El caso verifica edad menor a los 20 años"
            
            ElseIf (ThisWorkbook.Sheets("TZ13").Cells(i, 26).Value > 64) Then
            
                ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = "3"
                ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = "El caso verifica edad mayor a los 64 años"
               
            'verifico la antiguedad del resultado histologico/diagnostico
            ElseIf (ThisWorkbook.Sheets("TZ13").Cells(i, 27).Value > 1) Then
            
                ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = "3"
                ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = "La fecha del diagnóstico supera el año"
            
            
            'si esta dentro del rango verifica los datos obligatorios
            Else
                
                'depuro las variables
                leyenda = "Datos incompletos: "
                flag = 0
                
                For j = 18 To 22
                    
                    If (ThisWorkbook.Sheets("TZ13").Cells(i, j).Value = "" Or ThisWorkbook.Sheets("TZ13").Cells(i, j).Value = "No consta") Then

                        If (flag = 0) Then

                            leyenda = leyenda & ThisWorkbook.Sheets("TZ13").Cells(9, j).Value
                            flag = 1

                         Else

                            leyenda = leyenda & ", " & ThisWorkbook.Sheets("TZ13").Cells(9, j).Value

                        End If

                    End If

                Next


                If (flag = 1) Then
                    
                    ThisWorkbook.Sheets("TZ13").Cells(i, 28).Value = 3
                    ThisWorkbook.Sheets("TZ13").Cells(i, 30).Value = leyenda

                End If
            
                
            End If
            
        End If
        
    Else
    
        MsgBox ("TZ13. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub


'brief: analisa el formulario correspondiente al nombre y determina si los casos son validos o no
'param: void
'return: void


Sub analisis_tz14()

'declaracion de variables
Dim i As Integer
Dim j As Integer
Dim fecha As Date
Dim fecha2 As Date

'es la fila donde el auditor comienza a relevar
i = 10

'marca quien y cuando se analizo
ThisWorkbook.Sheets("tz14").Cells(5, 14).Value = Application.UserName
ThisWorkbook.Sheets("tz14").Cells(5, 15).Value = Date


'hace las siguientes lineas hasta que encuentra una celda vacia
Do Until ThisWorkbook.Sheets("tz14").Cells(i, 7).Value = ""
    
    'limpia la celda de categoria y fundamento por si se utiliza varias veces
    ThisWorkbook.Sheets("tz14").Cells(i, 14).Value = ""
    ThisWorkbook.Sheets("tz14").Cells(i, 16).Value = ""
    
    
    'para el motivo 3
    If (ThisWorkbook.Sheets("tz14").Cells(i, 7).Value = "Incompleto" Or ThisWorkbook.Sheets("tz14").Cells(i, 7).Value = "Completo") Then
        
        'verifica que la entrada al comite sea anterior a la fecha de obito
        If (ThisWorkbook.Sheets("tz14").Cells(i, 13).Value = "Si") Then
        
                ThisWorkbook.Sheets("tz14").Cells(i, 14).Value = 3
                ThisWorkbook.Sheets("tz14").Cells(i, 16).Value = "La fecha de entrada al comité es anterior a la fecha de obito"
        
        'verifica que la celda de diagnostico este vacia
        ElseIf (ThisWorkbook.Sheets("tz14").Cells(i, 10).Value = "") Then
        
            ThisWorkbook.Sheets("tz14").Cells(i, 14).Value = 3
            ThisWorkbook.Sheets("tz14").Cells(i, 16).Value = "No se verifica diagnóstico"
                
        'primero verifica si la fecha encontrada es distinta a la fecha envianda y luego verifica que se haya ingresado la fecha que se encontro
        ElseIf (ThisWorkbook.Sheets("tz14").Cells(i, 8).Value = "No" And ThisWorkbook.Sheets("tz14").Cells(i, 9).Value <> "") Then
            
            'paso la fecha encontrada a la variable
            fecha = ThisWorkbook.Sheets("tz14").Cells(i, 9).Value
            
            'verifico que la fecha encontrada este dentro del periodo evaluado con los datos en la hoja "Controles - Formulas"
            If (fecha > ThisWorkbook.Sheets("Controles - Formulas").Cells(14, 6).Value Or fecha < ThisWorkbook.Sheets("Controles - Formulas").Cells(15, 6).Value) Then

                ThisWorkbook.Sheets("tz14").Cells(i, 14).Value = 3
                ThisWorkbook.Sheets("tz14").Cells(i, 16).Value = "La fecha de la entrada al Comité no está dentro del período evaluado"

            End If
            
        End If
    
    Else
    
        MsgBox ("tz14. Hubo un error en la fila: " & i & ". Verificar con el che pibe")
    
    End If
    
    i = i + 1

Loop

End Sub

Sub analisis()

On Error Resume Next

    Call analisis_tz1
    Call analisis_tz2
    Call analisis_tz3
    Call analisis_tz4
    Call analisis_tz7
    Call analisis_tz8
    Call analisis_tz9
    Call analisis_tz10
    Call analisis_tz11
    Call analisis_tz12
    Call analisis_tz13
    Call analisis_tz14
    
On Error GoTo 0

End Sub
