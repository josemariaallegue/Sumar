
  Call cambioCampoCaracterANumerico(archivo, campo)
  
  
entrada = salida
salida = "para muestra CUIE x BENEF.IMD"
campo1 = "CUIE"
campo2 = "PROV_ID"
campo3 = ""
campo4 = ""
campo5 = ""
campo6 = ""
campo7 = ""
Call  Summarization(entrada, salida, campo1, campo2, campo3, campo4, campo5, campo6, campo7)



archivo = "para muestra CUIE x BENEF.IMD"
nombreViejo = "NUM_DE_REGS"
nombreNuevo ="CUIE_X_BENEF"
Call  cambioNombreCampoNumerico(archivo, nombreViejo, nombreNuevo)


entrada = "PARA MUESTRA.IMD"
'cruzar = "D:\SUMAR Crowe\Auxiliares\CUIE SELECCIONADOS.IMD" 
cruzar = "CUIE SELECCIONADOS.IMD" 
campo1 = "CUIE"
campo2 = "CUIE"
salida = "para muestra solo VALIDOS con CUIE select.IMD"
Call cruceTodosEnPrimaria(entrada, cruzar, campo1, campo2, salida)
  

  entrada = salida
salida = "para muestra solo VALIDOS con CUIE seleccionados.IMD"
condicion =  "  SELECCION <> """""
Call  exttracionConCondicion(entrada, salida, condicion)
  
  entrada = salida
salida = "seleccionados TOTAL.IMD"
campo1 = "SELECCION"
campo2 = "PROV_ID"
campo3 = ""
campo4 = ""
campo5 = ""
campo6 = ""
campo7 = ""
Call  Summarization(entrada, salida, campo1, campo2, campo3, campo4, campo5, campo6, campo7)

'GoTo fin

archivo = "seleccionados TOTAL.IMD"
nombreViejo = "NUM_DE_REGS"
nombreNuevo ="TOTAL_PROV"
Call  cambioNombreCampoNumerico(archivo, nombreViejo, nombreNuevo)

asdasd:

entrada = "para muestra solo VALIDOS con CUIE seleccionados.IMD"
cruzar = "seleccionados TOTAL.IMD"
campo1 = "SELECCION"
campo2 = "SELECCION"
salida = "para muestra cruce1.IMD"
Call cruceTodosEnPrimaria(entrada, cruzar, campo1, campo2, salida)
  
  entrada = salida
campo = "SELECCION1"
Call eliminarCampo(entrada , campo)

entrada = salida
cruzar = "para muestra CUIE x BENEF.IMD"
campo1 = "CUIE"
campo2 = "CUIE"
salida = "para muestra cruce2.IMD"
Call cruceTodosEnPrimaria(entrada, cruzar, campo1, campo2, salida)

  entrada = salida
campo = "CUIE1"
Call eliminarCampo(entrada , campo)
 
 
  entrada = salida
campo = "CUIE2"
Call eliminarCampo(entrada , campo)
 
 
'  entrada = salida
'cruzar = "D:\SUMAR Crowe\Auxiliares\Pagos - Codigos Elegibles.IMD" 
'cruzar = "CUIE SELECCIONADOS.IMD" 
'campo1 = "CODIGO_PRESTACION"
'campo2 = "CODIGOS_ELEGIBLES"
'salida = "para muestra cruce3.IMD"
'Call cruceTodosEnPrimaria(entrada, cruzar, campo1, campo2, salida)
  
 ' entrada = salida
'salida = "para muestra FINAL para excel.IMD"
'salida = "para muestra cruce4.IMD"
'condicion =  " CODIGOS_ELEGIBLES <> """""
'Call  exttracionConCondicion(entrada, salida, condicion)
crear = salida
campo = "EDAD"
decim = "0"
valor = "AÑOS_EN_DIA_PRESTACION"
Call agregoCampoNum(crear, campo, valor, decim)
crear = salida
campo = "PERTENECE_AL_FORMULARIO" 
longitud = 17
valor =  "CATEGORIA_LIQ" 
Call agregoCampoCaracter(crear, campo, longitud, valor)






  entrada = salida
salida = "prestaciones VALIDOS.IMD"
campo1 = "CUIE"
campo2 = "PROV_ID"
campo3 = ""
campo4 = ""
campo5 = ""
campo6 = ""
campo7 = ""
Call  Summarization(entrada, salida, campo1, campo2, campo3, campo4, campo5, campo6, campo7)

archivo = "prestaciones VALIDOS.IMD"
nombreViejo = "NUM_DE_REGS"
nombreNuevo ="CUIE_X_BENEF_VALIDOS"
Call  cambioNombreCampoNumerico(archivo, nombreViejo, nombreNuevo)

 entrada = "para muestra cruce2.IMD"
cruzar = "prestaciones VALIDOS.IMD"
campo1 = "CUIE"
campo2 = "CUIE"
salida = "para muestra FINAL para excel.IMD"
Call cruceTodosEnPrimaria(entrada, cruzar, campo1, campo2, salida)

  entrada = salida
campo = "CUIE1"
Call eliminarCampo(entrada , campo)


entrada = salida
salida = "para muestra FINAL para excel ordenado.IMD"
campo1 = "CUIE"
campo2 = "SELECCION"
Call ordenad2CamposAsc(entrada, salida, campo1, campo2)

crear = salida
decim = "2"
campo = "PROPORCION"
valor = "  @if(  CUIE_X_BENEF / TOTAL_PROV > 0,009;   CUIE_X_BENEF / TOTAL_PROV ; 0,01)"
Call agregoCampoNum(crear, campo, valor, decim)
campo = "N"
decim = "0"
valor = "TOTAL_PROV *0,5*(1-0,5)/((0,05^2/4)*(TOTAL_PROV -1)+0,5*(1-0,5))"
Call agregoCampoNum(crear, campo, valor, decim)
 campo = "CANTIDAD_MUESTRA"
valor = "    @if(    PROPORCION * N >= CUIE_X_BENEF ;   CUIE_X_BENEF; @if ( PROPORCION * N <6 ; 6; PROPORCION * N)  )"
Call agregoCampoNum(crear, campo, valor, decim)
 campo = "SEMILLA"
valor = "CUIE_X_BENEF_VALIDOS / CANTIDAD_MUESTRA"
Call agregoCampoNum(crear, campo, valor, decim)

crear = salida
campo = "CALCULOS" 
longitud = 100
valor =  """=SI(H"" +@AllTrim( @Str(  @Precno() + 1 ; 8;0) ) +  ""=H"" + @AllTrim(@Str( @Precno()  ; 8;0)) +  "";1+AM"" +@AllTrim( @Str(  @Precno() ; 8;0)) + "";1)""" 
Call agregoCampoCaracter(crear, campo, longitud, valor)


entrada = salida
salida = "para muestra FINAL.IMD"
Call extraxParaMuestraCampos(entrada, salida)


entrada = salida
condicion = ""
salida = "Muestra Pagos"
'Call ExpExcelConCondicion(entrada, condicion, salida, carpetaAñoMes)


fin:

End Sub


Function Sumarizacion(entrada, campo1, campo2, campo3, campo4, campo5, campo6, campo7, campo8, salida)
On Error GoTo nosumarizando3campos
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Summarization
	task.AddFieldToSummarize campo1
	If campo2 <> "" Then
	task.AddFieldToSummarize campo2
	End If
	If campo3 <> "" Then
	task.AddFieldToSummarize campo3
	End If
	If campo4 <> "" Then
	task.AddFieldToSummarize campo4
	End If
	If campo5 <> "" Then
	task.AddFieldToSummarize campo5
	End If
	If campo6 <> "" Then
	task.AddFieldToSummarize campo6
	End If
	If campo7 <> "" Then
	task.AddFieldToSummarize campo7
	End If
	If campo8 <> "" Then
	task.AddFieldToSummarize campo8
	End If

	dbName = salida
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_COUNT
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
nosumarizando3campos:
	Set task = Nothing
	Set db = Nothing

End Function


Function BusquedaDuplicados(entrada, campo1, campo2, campo3 , campo4, campo5, campo6, campo7, campo8, salida)
On Error GoTo nobusquedaRepetidosPor3Campos
	Set db = Client.OpenDatabase(entrada)
	Set task = db.DupKeyDetection
	task.IncludeAllFields
	task.AddKey campo1, "A"
	If campo2 <> "" Then
	task.AddKey campo2, "A"
	End If
	If campo3 <> "" Then
	task.AddKey campo3, "A"
	End If
		If campo4 <> "" Then
	task.AddKey campo4, "A"
	End If
		If campo5 <> "" Then
	task.AddKey campo5, "A"
	End If
		If campo6 <> "" Then
	task.AddKey campo6, "A"
	End If
		If campo7 <> "" Then
	task.AddKey campo7, "A"
	End If
		If campo8 <> "" Then
	task.AddKey campo8, "A"
	End If
		eqn = "" 
		task.Criteria = eqn 
	task.OutputDuplicates = TRUE
	dbName = salida
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
nobusquedaRepetidosPor3Campos:
End Function	


Function agregoCampoCaracter(crear, campo, longitud, valor)
On Error GoTo noagregoCampoCaracter
	Set db = Client.OpenDatabase(crear)
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
'	eqn = valor
	field.Name = campo
	field.Description = "" 
	field.Type = WI_VIRT_CHAR 
	field.Equation = valor
	field.Length = CLng(longitud)
	task.AppendField field 
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
	Client.CloseDatabase (crear)
noagregoCampoCaracter:
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
End Function

Function agregoCampoNum(crear, campo, valor, decim)
On Error GoTo noagregoCampoNum
	Set db = Client.OpenDatabase(crear)
	Set task = db.TableManagement
	Set table = db.TableDef
	Set field = table.NewField
'	eqn = "0"
	field.Name = campo
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = valor
	field.Decimals = CLng(decim)
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
	Client.CloseDatabase (crear)
noagregoCampoNum:
End Function


Function exttracionConCondicion(entrada, salida, condicion)
On Error GoTo noexttracionConCondicion
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = salida
	task.AddExtraction dbName, "", condicion
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
noexttracionConCondicion:
End Function

Function modificarCampos(archivo)
	Set db = Client.OpenDatabase(archivo)
	Set task = db.TableManagement
	Set table = db.TableDef
GoTo nohaaaa
	Set field = table.NewField
	field.Name = "FECHA_PRESTACION"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_PRESTACION", field
	Set field = table.NewField
	field.Name = "FECHA_NACIMIENTO"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_NACIMIENTO", field
	Set field = table.NewField  		' Caba lo na tiene
	field.Name = "FECHA_FACTURA"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_FACTURA", field
	Set field = table.NewField
	field.Name = "FECHA_RECEPCION"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_RECEPCION", field
	Set field = table.NewField
	field.Name = "FECHA_LIQUIDACION"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_LIQUIDACION", field
	Set field = table.NewField 		' Caba lo na tiene
	field.Name = "FECHA_DEBITO_BANCARIO"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_DEBITO_BANCARIO", field
	Set field = table.NewField 		' Caba lo na tiene
	field.Name = "FECHA_NOTIFICACION_PAGO"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "aaaa-mm-dd"
	task.ReplaceField "FECHA_NOTIFICACION_PAGO", field

	Set field = table.NewField
	field.Name = "IMPORTE_FACTURADO"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Decimals = 2
	field.IsImpliedDecimal = FALSE
	task.ReplaceField "IMPORTE_FACTURADO", field
	Set field = table.NewField
	field.Name = "IMPORTE_LIQUIDADO"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Decimals = 2
	field.IsImpliedDecimal = FALSE
	task.ReplaceField "IMPORTE_LIQUIDADO", field
	task.PerformTask
nohaaaa:
	Set task = Nothing
	Set db = Nothing
	Set table = Nothing
	Set field = Nothing
	Client.CloseDatabase (archivo)

End Function





Function cruceTodoEnPrimario(entrada, cruzar, salida, campo1, campo2)
On Error GoTo nocruceTodoEnPrimario
	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey campo1, campo2, "A"
	task.AddMatchKey "AÑO", "AÑO", "A"
	task.AddMatchKey "MES", "MES", "A"
	dbName = salida
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
nocruceTodoEnPrimario:	
	Set task = Nothing
	Set db = Nothing
End Function



Function Juntar31(arch1, arch2, arch3, arch4, arch5, arch6, arch7, arch8, arch9, arch10, arch11, arch12, arch13, arch14, arch15, arch16, arch17, arch18, arch19, arch20, arch21, arch22, arch23, arch24, arch25, arch26, arch27, arch28, arch29, arch30, arch31, crear) 
Exist1 = 0
Exist2 = 0
Exist3 = 0
Exist4 = 0
Exist5 = 0
Exist6 = 0
Exist7 = 0
Exist8 = 0
Exist9 = 0
Exist10 = 0
Exist11 = 0
Exist12 = 0
Exist13 = 0
Exist14 = 0
Exist15 = 0
Exist16 = 0
Exist17 = 0
Exist18 = 0
Exist19 = 0
Exist20 = 0
Exist21 = 0
Exist22 = 0
Exist23 = 0
Exist24 = 0
Exist25 = 0
Exist26 = 0
Exist27 = 0
Exist28 = 0
Exist29 = 0
Exist30 = 0
Exist31 = 0
	On Error GoTo fijar2
		Set dbObj = Client.OpenDatabase (arch1)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch1)
		If recordSet.count  > 0 Then  
			Exist1 = CLng(recordSet.count)
		Else 
			Exist1 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch1)
fijar2:
	On Error GoTo fijar3
		Set dbObj = Client.OpenDatabase (arch2)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch2)
		If recordSet.count  > 0 Then  
			Exist2 = CLng(recordSet.count)
		Else 
			Exist2 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch2)
fijar3:
	On Error GoTo fijar4
		Set dbObj = Client.OpenDatabase (arch3)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch3)
		If recordSet.count  > 0 Then  
			Exist3 = CLng(recordSet.count)
		Else 
			Exist3 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch3)
fijar4:
	On Error GoTo fijar5
		Set dbObj = Client.OpenDatabase (arch4)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch4)
		If recordSet.count  > 0 Then  
			Exist4 = CLng(recordSet.count)
		Else 
			Exist4 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch4)
fijar5:
	On Error GoTo fijar6
		Set dbObj = Client.OpenDatabase (arch5)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch5)
		If recordSet.count  > 0 Then  
			Exist5 = CLng(recordSet.count)
		Else 
			Exist5 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch5)
fijar6:
	On Error GoTo fijar7
		Set dbObj = Client.OpenDatabase (arch6)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch6)
		If recordSet.count  > 0 Then  
			Exist6 = CLng(recordSet.count)
		Else 
			Exist6 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch6)
fijar7:

	On Error GoTo fijar8
		Set dbObj = Client.OpenDatabase (arch7)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch7)
		If recordSet.count  > 0 Then  
			Exist7 = CLng(recordSet.count)
		Else 
			Exist7 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch7)
fijar8:
	On Error GoTo fijar9
		Set dbObj = Client.OpenDatabase (arch8)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch8)
		If recordSet.count  > 0 Then  
			Exist8 = CLng(recordSet.count)
		Else 
			Exist8 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch8)
fijar9:
	On Error GoTo fijar10
		Set dbObj = Client.OpenDatabase (arch9)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch9)
		If recordSet.count  > 0 Then  
			Exist9 = CLng(recordSet.count)
		Else 
			Exist9 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch9)
fijar10:
	On Error GoTo fijar11
		Set dbObj = Client.OpenDatabase (arch10)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch10)
		If recordSet.count  > 0 Then  
			Exist10 = CLng(recordSet.count)
		Else 
			Exist10 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch10)
fijar11:
	On Error GoTo fijar12
		Set dbObj = Client.OpenDatabase (arch11)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch11)
		If recordSet.count  > 0 Then  
			Exist11 = CLng(recordSet.count)
		Else 
			Exist11 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch11)
fijar12:
	On Error GoTo fijar13
		Set dbObj = Client.OpenDatabase (arch12)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch12)
		If recordSet.count  > 0 Then  
			Exist12 = CLng(recordSet.count)
		Else 
			Exist12 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch12)
fijar13:
	On Error GoTo fijar14
		Set dbObj = Client.OpenDatabase (arch13)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch13)
		If recordSet.count  > 0 Then  
			Exist13 = CLng(recordSet.count)
		Else 
			Exist13 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch13)
fijar14:
	On Error GoTo fijar15
		Set dbObj = Client.OpenDatabase (arch14)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch14)
		If recordSet.count  > 0 Then  
			Exist14 = CLng(recordSet.count)
		Else 
			Exist14 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch14)
fijar15:
	On Error GoTo fijar16
		Set dbObj = Client.OpenDatabase (arch15)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch15)
		If recordSet.count  > 0 Then  
			Exist15 = CLng(recordSet.count)
		Else 
			Exist15 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch15)
fijar16:

On Error GoTo fijar17
		Set dbObj = Client.OpenDatabase (arch16)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch16)
		If recordSet.count  > 0 Then  
			Exist16 = CLng(recordSet.count)
		Else 
			Exist16 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch16)
fijar17:

On Error GoTo fijar18
		Set dbObj = Client.OpenDatabase (arch17)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch17)
		If recordSet.count  > 0 Then  
			Exist17 = CLng(recordSet.count)
		Else 
			Exist17 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch17)
fijar18:

On Error GoTo fijar19
		Set dbObj = Client.OpenDatabase (arch18)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch18)
		If recordSet.count  > 0 Then  
			Exist18 = CLng(recordSet.count)
		Else 
			Exist18 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch18)
fijar19:

On Error GoTo fijar20
		Set dbObj = Client.OpenDatabase (arch19)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch19)
		If recordSet.count  > 0 Then  
			Exist19 = CLng(recordSet.count)
		Else 
			Exist19 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch19)
fijar20:

On Error GoTo fijar21
		Set dbObj = Client.OpenDatabase (arch20)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch20)
		If recordSet.count  > 0 Then  
			Exist20 = CLng(recordSet.count)
		Else 
			Exist20 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch20)
fijar21:

On Error GoTo fijar22
		Set dbObj = Client.OpenDatabase (arch21)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch21)
		If recordSet.count  > 0 Then  
			Exist21 = CLng(recordSet.count)
		Else 
			Exist21 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch21)
fijar22:

On Error GoTo fijar23
		Set dbObj = Client.OpenDatabase (arch22)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch22)
		If recordSet.count  > 0 Then  
			Exist22 = CLng(recordSet.count)
		Else 
			Exist22 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch22)
fijar23:

On Error GoTo fijar24
		Set dbObj = Client.OpenDatabase (arch23)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch23)
		If recordSet.count  > 0 Then  
			Exist23 = CLng(recordSet.count)
		Else 
			Exist23 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch23)
fijar24:

On Error GoTo fijar25
		Set dbObj = Client.OpenDatabase (arch24)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch24)
		If recordSet.count  > 0 Then  
			Exist24 = CLng(recordSet.count)
		Else 
			Exist24 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch24)
fijar25:

On Error GoTo fijar26
		Set dbObj = Client.OpenDatabase (arch25)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch25)
		If recordSet.count  > 0 Then  
			Exist25 = CLng(recordSet.count)
		Else 
			Exist25 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch25)
fijar26:

On Error GoTo fijar27
		Set dbObj = Client.OpenDatabase (arch26)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch26)
		If recordSet.count  > 0 Then  
			Exist26 = CLng(recordSet.count)
		Else 
			Exist26 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch26)
fijar27:

On Error GoTo fijar28
		Set dbObj = Client.OpenDatabase (arch27)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch27)
		If recordSet.count  > 0 Then  
			Exist27 = CLng(recordSet.count)
		Else 
			Exist27 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch27)
fijar28:

On Error GoTo fijar29
		Set dbObj = Client.OpenDatabase (arch28)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch28)
		If recordSet.count  > 0 Then  
			Exist28 = CLng(recordSet.count)
		Else 
			Exist28 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch28)
fijar29:

On Error GoTo fijar30
		Set dbObj = Client.OpenDatabase (arch29)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch29)
		If recordSet.count  > 0 Then  
			Exist29 = CLng(recordSet.count)
		Else 
			Exist29 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch29)
fijar30:

On Error GoTo fijar31
		Set dbObj = Client.OpenDatabase (arch30)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch30)
		If recordSet.count  > 0 Then  
			Exist30 = CLng(recordSet.count)
		Else 
			Exist30 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch30)
fijar31:
On Error GoTo fijar32
		Set dbObj = Client.OpenDatabase (arch31)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (arch31)
		If recordSet.count  > 0 Then  
			Exist31 = CLng(recordSet.count)
		Else 
			Exist31 = 0
		End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
	Client.CloseDatabase(arch30)
fijar32:
If Exist1 = 0  And Exist2 = 0 And Exist3 = 0 And Exist4 = 0 And Exist5 = 0 And Exist6 = 0 And Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0  And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then
	Set recordSet = Nothing
	Set db = Nothing
	Set dbObj = Nothing

GoTo noHayNada
end if

	Set recordSet = Nothing
	Set db = Nothing
	Set dbObj = Nothing

If Exist1 > 0 Then
	primero = arch1
	Exist1 = 0 
	If Exist2 = 0 And Exist3 = 0 And Exist4 = 0 And Exist5 = 0 And Exist6 = 0 And Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0  And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist2 > 0 Then
	primero = arch2
	Exist2 = 0
	If Exist3 = 0 And Exist4 = 0 And Exist5 = 0 And Exist6 = 0 And Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist3 > 0 Then
	primero = arch3
	Exist3 = 0
	If Exist4 = 0 And Exist5 = 0 And Exist6 = 0 And Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist4 > 0 Then
	primero = arch4
	Exist4 = 0
	If Exist5 = 0 And Exist6 = 0 And Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist5 > 0 Then
	primero = arch5
	Exist5 = 0
	If Exist6 = 0 And Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist6 > 0 Then
	primero = arch6
	Exist6 = 0
	If Exist7 = 0 And Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist7 > 0 Then
	primero = arch7
	Exist7 = 0
	If Exist8 = 0 And Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist8 > 0 Then
	primero = arch8
	Exist8 = 0
	If Exist9 = 0 And Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist9 > 0 Then
	primero = arch9
	Exist9 = 0
	If Exist10 = 0 And Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist10 > 0 Then
	primero = arch10
	Exist10 = 0
	If Exist11 = 0 And Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist11 > 0 Then
	primero = arch11
	Exist12 = 0
	If Exist12 = 0 And Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist12 > 0 Then
	primero = arch12
	Exist12 = 0
	If Exist13 = 0 And Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist13 > 0 Then
	primero = arch13
	Exist13 = 0
	If Exist14 = 0 And Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist14 > 0 Then
	primero = arch14
	Exist14 = 0
	If  Exist15 = 0 And Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist15 > 0 Then
	primero = arch15
	Exist15 = 0
	If  Exist16 = 0   And Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist16 > 0 Then
	primero = arch16
	Exist16 = 0
	If  Exist17 = 0   And Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist17> 0 Then
	primero = arch17
	Exist17 = 0
	If  Exist18 = 0   And Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist18 > 0 Then
	primero = arch18
	Exist18 = 0
	If  Exist19 = 0   And Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist19 > 0 Then
	primero = arch19
	Exist19 = 0
	If  Exist20 = 0   And Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist20 > 0 Then
	primero = arch20
	Exist20 = 0
	If Exist21 = 0   And Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist21 > 0 Then
	primero = arch21
	Exist21 = 0
	If  Exist22 = 0   And Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist22 > 0 Then
	primero = arch22
	Exist22 = 0
	If Exist23 = 0   And Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist23 > 0 Then
	primero = arch23
	Exist23 = 0
	If Exist24 = 0 And Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist24 > 0 Then
	primero = arch24
	Exist24 = 0
	If Exist25 = 0 And Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist25 > 0 Then
	primero = arch25
	Exist25 = 0
	If Exist26 = 0 And Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist26 > 0 Then
	primero = arch26
	Exist26 = 0
	If Exist27 = 0 And Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist27 > 0 Then
	primero = arch27
	Exist27 = 0
	If Exist28 = 0  And Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist28 > 0 Then
	primero = arch28
	Exist28 = 0
	If Exist29 = 0 And Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist29 > 0 Then
	primero = arch29
	Exist29 = 0
	If Exist30 = 0 And Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist30 > 0 Then
	primero = arch30
	Exist30 = 0
	If Exist31 = 0 Then 
		GoTo extraxion
	End If
ElseIf Exist31 > 0 Then
	primero = arch31
		GoTo extraxion
Else
GoTo noHayNada
End If

	Set db = Client.OpenDatabase(primero)
	Set task = db.AppendDatabase 
	If  exist2 > 0 Then 
	task.AddDatabase arch2
	End If 
	If  exist3 > 0 Then 
	task.AddDatabase arch3
	End If 
	If  exist4 > 0 Then 
	task.AddDatabase arch4
	End If 
	If  exist5 > 0 Then 
	task.AddDatabase arch5
	End If 
	If  exist6 > 0 Then 
	task.AddDatabase arch6
	End If
	If  exist7 > 0 Then 
	task.AddDatabase arch7
	End If
	If  exist8 > 0 Then 
	task.AddDatabase arch8
	End If 
	If  exist9 > 0 Then 
	task.AddDatabase arch9
	End If
	If  exist10 > 0 Then 
	task.AddDatabase arch10
	End If
	If  exist11 > 0 Then 
	task.AddDatabase arch11
	End If
	If  exist12 > 0 Then 
	task.AddDatabase arch12
	End If
	If  exist13 > 0 Then 
	task.AddDatabase arch13
	End If
	If  exist14 > 0 Then 
	task.AddDatabase arch14
	End If 
	If  exist15 > 0 Then 
	task.AddDatabase arch15
	End If
	If  exist16 > 0 Then 
	task.AddDatabase arch16
	End If
	If  exist17 > 0 Then 
	task.AddDatabase arch17
	End If
	If  exist18 > 0 Then 
	task.AddDatabase arch18
	End If
	If  exist19 > 0 Then 
	task.AddDatabase arch19
	End If
	If  exist20 > 0 Then 
	task.AddDatabase arch20
	End If
	If  exist21 > 0 Then 
	task.AddDatabase arch21
	End If
	If  exist22 > 0 Then 
	task.AddDatabase arch22
	End If
	If  exist23 > 0 Then 
	task.AddDatabase arch23
	End If
	If  exist24 > 0 Then 
	task.AddDatabase arch24
	End If
	If  exist25 > 0 Then 
	task.AddDatabase arch25
	End If
	If  exist26 > 0 Then 
	task.AddDatabase arch26
	End If
	If  exist27 > 0 Then 
	task.AddDatabase arch27
	End If
	If  exist28 > 0 Then 
	task.AddDatabase arch28
	End If
	If  exist29 > 0 Then 
	task.AddDatabase arch29
	End If
	If  exist30 > 0 Then 
	task.AddDatabase arch30
	End If
	If  exist31 > 0 Then 
	task.AddDatabase arch31
	End If
	task.PerformTask crear, ""
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (primero)
	GoTo noHayNada
extraxion:
On Error GoTo noHayNada
		Set db = Client.OpenDatabase(primero)
		Set task = db.Extraction
		task.IncludeAllFields
		dbName = crear
		task.AddExtraction dbName, "", condicion
		task.PerformTask 1, db.Count
		Set task = Nothing
		Set db = Nothing
		Client.CloseDatabase (primero)
noHayNada:
		Set task = Nothing
		Set db = Nothing
End Function


Function Summarization(entrada, salida, campo1, campo2, campo3, campo4, campo5, campo6, campo7)
On Error GoTo noSummarization
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Summarization
	task.AddFieldToSummarize campo1
	If campo2 <> "" Then
	task.AddFieldToSummarize campo2
	End If 
	If campo3 <> "" Then
	task.AddFieldToSummarize campo3
	End If 
		If campo4 <> "" Then
	task.AddFieldToSummarize campo4
	End If 
		If campo5 <> "" Then
	task.AddFieldToSummarize campo5
	End If 
		If campo6 <> "" Then
	task.AddFieldToSummarize campo6
	End If 
		If campo7 <> "" Then
	task.AddFieldToSummarize campo7
	End If 
	dbName = salida
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_COUNT
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
noSummarization:
End Function



Function cruzeSinCoincidencia(entrada, cruzar, campo1, campo2, resultado)
On Error GoTo nocruzeSinCoincidencia
		Set dbObj = Client.OpenDatabase (cruzar)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (cruzar)
		If recordSet.count  > 0 Then  
		Client.CloseDatabase (cruzar)
		Set dbObj = Nothing
	 	Set recordSet = Nothing
	 	Set db = Nothing
	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.AddMatchKey campo1, campo2, "A"
	task.AddMatchKey "AÑO", "AÑO", "A"
	task.AddMatchKey "MES", "MES", "A"
	dbName = resultado
	task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (archivo)
		Else
		Client.CloseDatabase (cruzar)
		Set dbObj = Nothing
	 	Set recordSet = Nothing
	 	Set db = Nothing
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = resultado
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)

	
	
		End If
	
nocruzeSinCoincidencia:
End Function



Function cruzeSinCoincidencia1(entrada, cruzar, campo1, campo2, resultado)
On Error GoTo nocruzeSinCoincidencia1
		Set dbObj = Client.OpenDatabase (cruzar)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (cruzar)
		If recordSet.count  > 0 Then  
		Client.CloseDatabase (cruzar)
		Set dbObj = Nothing
	 	Set recordSet = Nothing
	 	Set db = Nothing

	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.AddMatchKey campo1, campo2, "A"
	dbName = resultado
	task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (archivo)
		Else
		Client.CloseDatabase (cruzar)
		Set dbObj = Nothing
	 	Set recordSet = Nothing
	 	Set db = Nothing

	Set db = Client.OpenDatabase(entrada)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = resultado
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)

	
	
		End If
	
nocruzeSinCoincidencia1:
End Function




Function eliminarCampo(entrada , campo)
On Error GoTo noeliminarCampo
	Set db = Client.OpenDatabase(entrada )
	Set task = db.TableManagement
	task.RemoveField campo
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (archivo)
noeliminarCampo:
	Set task = Nothing
	Set db = Nothing
End Function




Function agregoABase(primero, segundo, campo, campo2, salida)
On Error GoTo noagregoABase
		Set dbObj = Client.OpenDatabase (segundo)  
	 	Set recordSet = dbObj.RecordSet  
	 	Set db = Client.OpenDatabase (segundo)
		If recordSet.count  > 0 Then  
			Set db = Client.OpenDatabase(primero)
			Set task = db.JoinDatabase
			task.FileToJoin segundo
			task.IncludeAllPFields
			task.AddSFieldToInc campo
			If campo2  <> "" Then
			task.AddSFieldToInc campo2
			End If 
			task.AddMatchKey "ID", "ID", "A"
			dbName = salida
			task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
			Set task = Nothing
			Set db = Nothing
			Client.CloseDatabase (primero)	
		Else 
		
			Set db = Client.OpenDatabase(primero)
			Set task = db.TableManagement
			Set table = db.TableDef
			Set field = table.NewField
			eqn = """"""
			field.Name = campo
			field.Description = "" 
			field.Type = WI_VIRT_CHAR 
			field.Equation = eqn
			field.Length = 1
			task.AppendField field 
			task.PerformTask
			Set task = Nothing
			Set db = Nothing
			Set table = Nothing
			Set field = Nothing
			
			'If campo2  <> "" Then
			Set db = Client.OpenDatabase(primero)
			Set task = db.TableManagement
			Set table = db.TableDef
			Set field = table.NewField
			eqn = """"""
			field.Name = campo2
			field.Description = "" 
			field.Type = WI_VIRT_CHAR 
			field.Equation = eqn
			field.Length = 1
			task.AppendField field 
			task.PerformTask
			Set task = Nothing
			Set db = Nothing
			Set table = Nothing
			Set field = Nothing

		'End If 

			
			
			Client.CloseDatabase (primero)
			
	Set db = Client.OpenDatabase(primero)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = salida
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (primero)
			
End If
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
		Client.CloseDatabase(segundo)
GoTo nononoAgregoABase	
noagregoABase:



			Set db = Client.OpenDatabase(primero)
			Set task = db.TableManagement
			Set table = db.TableDef
			Set field = table.NewField
			eqn = """"""
			field.Name = campo
			field.Description = "" 
			field.Type = WI_VIRT_CHAR 
			field.Equation = eqn
			field.Length = 1
			task.AppendField field 
			task.PerformTask
			Set task = Nothing
			Set db = Nothing
			Set table = Nothing
			Set field = Nothing
			
			Set db = Client.OpenDatabase(primero)
			Set task = db.TableManagement
			Set table = db.TableDef
			Set field = table.NewField
			eqn = """"""
			field.Name = campo2
			field.Description = "" 
			field.Type = WI_VIRT_CHAR 
			field.Equation = eqn
			field.Length = 1
			task.AppendField field 
			task.PerformTask
			Set task = Nothing
			Set db = Nothing
			Set table = Nothing
			Set field = Nothing
			
	Set db = Client.OpenDatabase(primero)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = salida
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (primero)

nononoAgregoABase:
		Set recordSet = Nothing
		Set db = Nothing
		Set dbObj = Nothing
End Function



Function busquedaRepetidosResumen(entrada, cruzar, resultado, campo)
On Error GoTo nobusquedaRepetidosResumen
	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.AddSFieldToInc campo
	task.AddMatchKey "NRO_DOC", "NRO_DOC", "A"
	task.AddMatchKey "CODIGO_PRESTACION", "CODIGO_PRESTACION", "A"
	task.AddMatchKey "CUIE_EFECTOR", "CUIE_EFECTOR", "A"
	task.AddMatchKey "FECHA_PRESTACION", "FECHA_PRESTACION", "A"
	dbName = resultado
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
nobusquedaRepetidosResumen:
End Function


Function busquedaRepetidosResumenB(entrada, cruzar, resultado, campo)
On Error GoTo nobusquedaRepetidosResumenB
	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.AddSFieldToInc campo
	task.AddMatchKey "NRO_DOC", "NRO_DOC", "A"
	task.AddMatchKey "CODIGO_PRESTACION", "CODIGO_PRESTACION", "A"
'	task.AddMatchKey "CUIE_EFECTOR", "CUIE_EFECTOR", "A"
'	task.AddMatchKey "FECHA_PRESTACION", "FECHA_PRESTACION", "A"
	dbName = resultado
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
nobusquedaRepetidosResumenB:
End Function



Function cruceTodosEnPrimaria(entrada, cruzar, campo1, campo2, salida)
On Error GoTo nocruceTodosEnPrimaria
	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey campo1, campo2, "A"
	task.AddMatchKey "PROV_ID", "PROV_ID", "A"

	dbName = salida
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
nocruceTodosEnPrimaria:
End Function


Function crucePrestacionesIncompatibles(entrada, cruzar, campo1, campo2, salida)
On Error GoTo noCruce
	Set db = Client.OpenDatabase(entrada)
	Set task = db.JoinDatabase
	task.FileToJoin cruzar
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey campo1, campo2, "A"
	task.Criteria = ""
	task.CreateVirtualDatabase = False
	dbName = salida
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)	
noCruce:
	Set task = Nothing
	Set db = Nothing
End Function

 
Function cambioNombreCampoCaracter(archivo, campoViejo, campoNuevo)	
	Set db = Client.OpenDatabase(archivo)
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = campoNuevo
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 6
	task.ReplaceField campoViejo, field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function




Function busquedaPorPrimerNombre(primero, segundo, campo, resultado)
On Error GoTo nobusquedaPorPrimerNombre
	Set db = Client.OpenDatabase(primero)
	Set task = db.JoinDatabase
	task.FileToJoin segundo
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey campo, "NOMBRE", "A"
	dbName = resultado
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (primero)
nobusquedaPorPrimerNombre:
End Function


Function busquedaSinSegundoNombre(primero, segundo, campo, resultado)
On Error GoTo nobusquedaSinSegundoNombre
	Set db = Client.OpenDatabase(primero)
	Set task = db.JoinDatabase
	task.FileToJoin segundo
	task.IncludeAllPFields
	task.AddMatchKey campo, "NOMBRE", "A"
	dbName = resultado
	task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (primero)
nobusquedaSinSegundoNombre:
End Function


Function agrego1ArchivoAlOtro(primero, segundo, resultado)
On Error GoTo noagrego1ArchivoAlOtro
	Set db = Client.OpenDatabase(primero)
	Set task = db.AppendDatabase
	task.AddDatabase segundo
	dbName = resultado
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (primero)
noagrego1ArchivoAlOtro:
End Function



Function encontradosPor2NombresSumarizacion(archivo, resultado)
On Error GoTo noencontradosPor2NombresSumarizacion
	Set db = Client.OpenDatabase(archivo)
	Set task = db.Summarization
	task.AddFieldToSummarize "ID"
	task.AddFieldToSummarize "APELLIDO_BENEFICIARIO"
	task.AddFieldToSummarize "NOMBRE_BENEFICIARIO"
	task.AddFieldToInc "PRIMER_NOMBRE"
	task.AddFieldToInc "SEGUNDO_NOMBRE"
	task.AddFieldToInc "NOMBRE_MINUSCULA"
	task.AddFieldToInc "SEXO"
	task.AddFieldToInc "CIRCULAR"
	task.AddFieldToInc "ORÍGEN_Y_COMENTARIOS"
	task.AddFieldToInc "NOMBRE"
	dbName = resultado
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_COUNT
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (archivo)
noencontradosPor2NombresSumarizacion:
End Function


Function cambioNombreCampoNumerico(archivo, nombreViejo, nombreNuevo)
	Set db = Client.OpenDatabase(archivo)
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = nombreNuevo
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField nombreViejo, field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function


Function ordenad2CamposAsc(entrada, salida, campo1, campo2)
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Sort
	task.AddKey campo1, "A"
	task.AddKey campo2, "A"
	dbName = salida
	task.PerformTask dbName
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
End Function




Function ExpExcelConCondicion(entrada, condicion, salida, carpetaAñoMes)
On Error GoTo noExpExcelConCondicion
	Set db = Client.OpenDatabase(entrada)
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = condicion
	task.PerformTask "D:\SUMAR Crowe\Exportaciones\"+ salida +".XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
	Client.CloseDatabase (entrada)
noExpExcelConCondicion:
End Function



Function extraxParaMuestraCampos(entrada, salida)
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Extraction
task.AddFieldToInc "CLAVE_BENEFICIARIO"
	task.AddFieldToInc "BENEF_APELLIDO"
	task.AddFieldToInc "BENEF_NOMBRE"
	task.AddFieldToInc "BENEF_TIPO_DOC_ORIGINAL"
	task.AddFieldToInc "BENEF_NRO_DOC_ORIGINAL"
	task.AddFieldToInc "BENEF_FECHA_NACIMIENTO"
	task.AddFieldToInc "CEB"
	task.AddFieldToInc "CUIE"
	task.AddFieldToInc "FECHA_ULTIMA_PRESTACION"
	task.AddFieldToInc "CODIGO_PRESTACION"
	task.AddFieldToInc "DEVENGA_CAPITA"
	task.AddFieldToInc "DEVENGA_CANTIDAD_CAPITA"
	task.AddFieldToInc "AÑOS_EN_DIA_PRESTACION_MIN"
	task.AddFieldToInc "GRUPO_POBLACIONAL_MIN"
	task.AddFieldToInc "AÑOS_EN_DIA_PRESTACION"
	task.AddFieldToInc "GRUPO_POBLACIONAL"
	task.AddFieldToInc "AÑOS_EN_DIA_PRESTACION_MAX"
	task.AddFieldToInc "GRUPO_POBLACIONAL_MAX"
	task.AddFieldToInc "ESTADO_AUDITORIA"
	task.AddFieldToInc "MOTIVO_AUDITORIA"
	task.AddFieldToInc "SITUACION"
	task.AddFieldToInc "UN_SOLO_MOTIVO"
	task.AddFieldToInc "PROV_ID"
	task.AddFieldToInc "CENTRO"
	task.AddFieldToInc "DENOMINACIONLEGAL"
	task.AddFieldToInc "DIRECCION"
	task.AddFieldToInc "LOCALIDAD"
	task.AddFieldToInc "PROVINCIA"
	task.AddFieldToInc "CATEGORIA_LIQUIDACION"
	task.AddFieldToInc "GRUPO_LIQUIDACION"
	task.AddFieldToInc "TOTAL_PROV"
	task.AddFieldToInc "CUIE_X_BENEF"
	task.AddFieldToInc "EDAD"
	task.AddFieldToInc "CUIE_X_BENEF_VALIDOS"
	task.AddFieldToInc "PROPORCION"
	task.AddFieldToInc "N"
	task.AddFieldToInc "CANTIDAD_MUESTRA"
	task.AddFieldToInc "SEMILLA"
	task.AddFieldToInc "CALCULOS"

	dbName = salida
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function



	
Function extraxionMustraSOLO(entrada, salida)
	Set db = Client.OpenDatabase(entrada)
	Set task = db.Extraction
	task.AddFieldToInc "CLAVE_BENEFICIARIO"
	task.AddFieldToInc "BENEF_APELLIDO"
	task.AddFieldToInc "BENEF_NOMBRE"
	task.AddFieldToInc "BENEF_TIPO_DOC_ORIGINAL"
	task.AddFieldToInc "BENEF_NRO_DOC_ORIGINAL"
	task.AddFieldToInc "BENEF_FECHA_NACIMIENTO"
	task.AddFieldToInc "CEB"
	task.AddFieldToInc "CUIE"
	task.AddFieldToInc "FECHA_ULTIMA_PRESTACION"
	task.AddFieldToInc "CODIGO_PRESTACION"
	task.AddFieldToInc "DEVENGA_CAPITA"
	task.AddFieldToInc "DEVENGA_CANTIDAD_CAPITA"
	task.AddFieldToInc "AÑOS_EN_DIA_PRESTACION_MIN"
	task.AddFieldToInc "GRUPO_POBLACIONAL_MIN"
	task.AddFieldToInc "AÑOS_EN_DIA_PRESTACION"
	task.AddFieldToInc "GRUPO_POBLACIONAL"
	task.AddFieldToInc "AÑOS_EN_DIA_PRESTACION_MAX"
	task.AddFieldToInc "GRUPO_POBLACIONAL_MAX"
	task.AddFieldToInc "ESTADO_AUDITORIA"
	task.AddFieldToInc "MOTIVO_AUDITORIA"
'	task.AddFieldToInc "ESTADO_AUDITORIA_DICIEMBRE"
'	task.AddFieldToInc "EN_DICIEMBRE_2012"
 	task.AddFieldToInc "SITUACION"
	task.AddFieldToInc "UN_SOLO_MOTIVO"
	task.AddFieldToInc "PROV_ID"
	task.AddFieldToInc "CENTRO"
	task.AddFieldToInc "DENOMINACIONLEGAL"
	task.AddFieldToInc "DIRECCION"
	task.AddFieldToInc "LOCALIDAD"
	task.AddFieldToInc "PROVINCIA"
	task.AddFieldToInc "CATEGORIA_LIQUIDACION"
	task.AddFieldToInc "GRUPO_LIQUIDACION"
	dbName = salida
	task.AddExtraction dbName, "", ""
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.CloseDatabase (entrada)
End Function


Function cambioCampoCaracterANumerico(archivo, campo)
	Set db = Client.OpenDatabase("PARA MUESTRA.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PROV_ID"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "PROV_ID", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function