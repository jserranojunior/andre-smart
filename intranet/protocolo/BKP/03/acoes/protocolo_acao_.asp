<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
   <!--include file="../includes/inc_area_restrita.asp"-->
<%
'***** Variáveis ******

data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")

acao = request("acao") 
obj_excl = request("obj_excl")

cd_unidade = request("cd_unidade")
cd_protocolo = request("cd_protocolo")
cd_digito = request("cd_digito")
cd_codigo = request("cd_codigo")
cd_material = request("cd_material")

modo = request("modo")
tipo = request("tipo")

nm_nome = request("nm_nome")
cd_idade = request("cd_idade")
nm_registro = request("nm_registro")
cd_convenio = request("cd_convenio")
	cd_convenio_1 = request("cd_convenio_1")
nm_sexo = request("nm_sexo")
nm_cirug_realizada = request("nm_cirug_realizada")

	dt_dia_cirurgia = request("dt_dia_cirurgia")
	dt_mes_cirurgia = request("dt_mes_cirurgia")
	dt_ano_cirurgia = request("dt_ano_cirurgia")
dt_cirurgia = dt_mes_cirurgia&"/"&dt_dia_cirurgia&"/"&dt_ano_cirurgia

nm_agenda_cirurgia = request("nm_agenda_cirurgia") 

	dt_hora_inicio = request("dt_hora_inicio")
	dt_minuto_inicio = request("dt_minuto_inicio")
dt_hr_inicio = dt_hora_inicio&":"&dt_minuto_inicio

	dt_hora_fim = request("dt_hora_fim")
	dt_minuto_fim = request("dt_minuto_fim")
dt_hr_fim = dt_hora_fim&":"&dt_minuto_fim

	dt_hora_agenda = request("dt_hora_agenda")
	dt_minuto_agenda = request("dt_minuto_agenda")
	
		if dt_hora_agenda <> "" AND dt_minuto_agenda <> "" Then
			dt_hr_agenda = dt_hora_agenda&":"&dt_minuto_agenda&":00"
		else
			dt_hr_agenda = dt_hr_inicio
		end if

	cd_crm_1 = request("cd_crm_1")
	if cd_user = 46 then
	response.write("***( "&cd_crm_1&" )***")
	end if
	
cd_rack = request("cd_rack")
	cd_rack_1 = request("cd_rack_1")

cd_especialidade = request("cd_especialidade")
	cd_especialidade_1 = request("cd_especialidade_1")
	
cd_co = request("cd_co")

procedimentos_total = request("procedimentos_total")
	if right(procedimentos_total,1) = "," Then
	proc_len = len(procedimentos_total)
	proc_len = proc_len - 1
	
		procedimentos_total = Left(procedimentos_total,proc_len)
	end if

materiais_total = request("materiais_total")
	if right(materiais_total,1) = "," Then
	mat_len = len(materiais_total)
	mat_len = mat_len - 1
	
		materiais = Left(materiais_total,mat_len)
	end if
		if materiais_total <> "" Then ' Garante que o protocolo será classificado corretamente em caso de empréstimo
			nm_agenda_cirurgia = "E" 
		end if

subtotal_patrimonio = request("subtotal_patrimonio")
	if right(subtotal_patrimonio,1) = "," Then
	pat_len = len(subtotal_patrimonio)
	pat_len = pat_len - 1
	
		Patrimonios = Left(subtotal_patrimonio,pat_len)
	end if

'************ Condições ************************
if cd_convenio_1 <> "" Then
	cd_convenio = cd_convenio_1
end if

if cd_rack_1 <> "" Then
	cd_rack = cd_rack_1
end if
	if cd_rack = "" Then
	cd_rack = "11"
	end if

if cd_especialidade_1 <> "" Then
	cd_especialidade = cd_especialidade_1
end if

if dt_hr_agenda = ":" Then
	dt_hr_agenda = NULL
End if

if dt_hr_inicio = ":" Then
	dt_hr_inicio = NULL
End if

if dt_hr_fim = ":" Then
	dt_hr_fim = NULL
End if

'***** verifica se já existe no Banco de Dados *****
		'xsql ="SELECT cd_protocolo FROM TBL_PROTOCOLO WHERE cd_protocolo = '"&cd_protocolo&"'"
'		xsql ="SELECT A002_numpro FROM TBL_PROTOCOLO WHERE A002_numpro = '"&cd_protocolo&"'"
'		Set rs = dbconn.execute(xsql)
'			Do while NOT rs.EOF
'			cd_prot = rs("A002_numpro")
'			rs.movenext
'			Loop	

'if cd_prot = "" Then
	
	'***** Ações ******
x = 100
	
if x = 100 then	'evita que os dados sejam gravados no Banco de dados, fins de desenvolvimento



'***************************************************************************************************************************************************************
'*** INSERIR PROTOCOLO ****************************************************************************************************************************************
'*************************************************************************************************************************************************************
	If acao = "inserir" Then

	xsql ="up_protocolo_insert @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm_1&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"',@cd_co='"&cd_co&"'"
	dbconn.execute(xsql)
	response.write("OK, inserido")
	mensagem = "OK, protocolo inserido com sucesso"
'	response.redirect "../protocolo.asp"
		'Verifica os procedimentos
		
		'*** Procedimentos ***
		'++++++++++++++++++++++++++++
		'+++ INSERE PROCEDIMENTOS +++
		'++++++++++++++++++++++++++++
		numero = 1
		proc_array = split(procedimentos_total,",")
		For Each proc_item In proc_array
			
				if numero = 1 then
				proc_prio = "S"
				Else
				proc_prio = "N"
				End if
			'Response.Write("<br>***" & proc_item)
			
		if IsNumeric(proc_item) then
			xsql ="SELECT * FROM TBL_protocolo_procedimento WHERE A002_numpro ='"&cd_protocolo&"' AND A053_codung = '"&cd_unidade&"' AND A058_codpro ='"&proc_item&"'"
			Set rs_proc = dbconn.execute(xsql)
				
				if not rs_proc.EOF then
					'Altera o Procedimento
					cd_prot_proc = rs_proc("cd_codigo")
					xsql ="up_protocolo_procedimento_update @cd_prot_proc = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @procedimento='"&proc_item&"', @primeira='"&proc_prio&"'"
					dbconn.execute(xsql)
					'response.write(" "&numero&" existe")
					mensagem1 = "Procedimento já existe."
				elseif rs_proc.EOF Then
					'Insere o procedimento
					xsql_proc ="up_protocolo_procedimento_insert @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @procedimento='"&proc_item&"', @primeira='"&proc_prio&"'"
					dbconn.execute(xsql_proc)
					'response.write(" "&numero&" não existe")
					mensagem2 = "Procedimento inserido."				
				End if
		end if
				
			numero = numero + 1
		Next
	
	
		'*** Materiais *** 
		'++++++++++++++++++++++++
		'+++ INSERE MATERIAIS +++
		'++++++++++++++++++++++++
	if materiais_total <> "" then
	numero = 1
		mat_array = split(materiais_total,",")
		For Each mat_item In mat_array
		
				'separador = Instr(mat_item,";")
				'mat_item_qtd = replace(right(mat_item,(separador)),";","")
				'	response.write(mat_item_qtd)
				'mat_item = replace(left(mat_item,(separador)),";","")
				'	response.write(mat_item)
					
				'Quantidade de caracteres
					mat_caracteres = len(mat_item)
				'Posição do separador
					separador = Instr(mat_item,";")
				'Quantidade de materiais
					mat_item_qtd = mid(mat_item,separador+1,mat_caracteres)
				'Código do material
					mat_item = mid(mat_item,1,separador-1)
			
		if IsNumeric(mat_item) then
			Response.Write("<br>" & mat_item)
			xsql ="SELECT * FROM TBL_protocolo_material WHERE A061_codmat ='"&cd_protocolo&"' AND A053_codung = '"&cd_unidade&"' AND A061_codmat ='"&mat_item&"'"
			Set rs_mat = dbconn.execute(xsql)
			
				
				'if not rs_mat.EOF then
					'Altera o Material
				'	cd_prot_mat = rs_mat("cd_codigo")
				'	xsql ="up_protocolo_material @cd_prot_mat = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @material='"&mat_item&"', @primeira=N"
				'	dbconn.execute(xsql)
				'	'response.write(" "&numero&" existe")
				'	mensagem3 = "Material já existe."
				'else
				if rs_mat.EOF Then
					'Insere o Material
					xsql_mat ="up_protocolo_material_insert @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @material='"&mat_item&"', @quantidade='"&mat_item_qtd&"'"
					dbconn.execute(xsql_mat)
					'response.write(" "&numero&" não existe")
					mensagem4 = "material inserido."
				End if
		end if
			numero = numero + 1
		Next
	end if

	
		'*** Patrimonios *** 
		'++++++++++++++++++++++++++
		'+++ INSERE PATRIMONIOS +++
		'++++++++++++++++++++++++++
		numero = 1
		pat_array = split(subtotal_patrimonio,",")
		For Each pat_item In pat_array
			
		if IsNumeric(pat_item) then
			xsql ="SELECT * FROM TBL_protocolo_patrimonio WHERE cd_protocolo ='"&cd_protocolo&"' AND cd_unidade = '"&cd_unidade&"' AND cd_patrimonio ='"&pat_item&"'"
			Set rs_pat = dbconn.execute(xsql)
				
				if not rs_proc.EOF then
					'Altera o Procedimento
					cd_prot_pat = rs_pat("cd_codigo")
					xsql ="up_protocolo_patrimonio_update @cd_prot_pat = '"&cd_prot_pat&"', @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @patrimonio='"&pat_item&"'"
					dbconn.execute(xsql)
					''response.write(" "&numero&" existe")
					mensagem1 = "Patrimonio já existe."
				elseif rs_pat.EOF Then
					'Insere o procedimento
					xsql_pat ="up_protocolo_patrimonio_insert @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @patrimonio='"&pat_item&"'"
					dbconn.execute(xsql_pat)
					'response.write(" "&numero&" não existe")
					mensagem2 = "Patrimonio inserido."				
				End if
		end if
				
			numero = numero + 1
		Next
	
	
	Elseif acao = "editar" Then
'***************************************************************************************************************************************************************
'*** ALTERAR PROTOCOLO ****************************************************************************************************************************************
'*************************************************************************************************************************************************************
	
	xsql ="up_protocolo_update @cd_codigo='"&cd_codigo&"', @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm_1&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"',@cd_co='"&cd_co&"'"
	dbconn.execute(xsql)
	response.write("OK, alterado")
	mensagem = "Ok, protocolo alterado com sucesso"
		'*** Procedimentos ***
		'::::::::::::::::::::::::::::
		'::: INSERE PROCEDIMENTOS :::
		'::::::::::::::::::::::::::::
		numero = 1
		proc_array = split(procedimentos_total,",")
		For Each proc_item In proc_array
			
				if numero = 1 then
				proc_prio = "S"
				Else
				proc_prio = "N"
				End if
			'Response.Write("<br>***" & proc_item)
			
		if IsNumeric(proc_item) then
			xsql ="SELECT * FROM TBL_protocolo_procedimento WHERE A002_numpro ='"&cd_protocolo&"'AND a053_codung = '"&cd_unidade&"' AND A058_codpro ='"&proc_item&"'"
			Set rs_proc = dbconn.execute(xsql)
				
				if not rs_proc.EOF then
					'Altera o Procedimento
					cd_prot_proc = rs_proc("cd_codigo")
					xsql ="up_protocolo_procedimento_update @cd_prot_proc = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @unidade = '"&cd_unidade&"', @procedimento='"&proc_item&"', @primeira='"&proc_prio&"'"
					dbconn.execute(xsql)
					'response.write(" "&numero&" existe")
					mensagem1 = "Procedimento já existe."
				elseif rs_proc.EOF Then
					'Insere o procedimento
					xsql_proc ="up_protocolo_procedimento_insert @protocolo='"&cd_protocolo&"', @unidade = '"&cd_unidade&"', @procedimento='"&proc_item&"', @primeira='"&proc_prio&"'"
					dbconn.execute(xsql_proc)
					'response.write(" "&numero&" não existe")
					mensagem2 = "Procedimento inserido."
				End if
		end if
				
			numero = numero + 1
		Next
	
	
		'*** Materiais *** 
		'::::::::::::::::::::::::
		'::: INSERE Materiais :::
		'::::::::::::::::::::::::
			if materiais_total <> "" then
			numero = 1
				mat_array = split(materiais_total,",")
				For Each mat_item In mat_array
				
						'Quantidade de caracteres
							mat_caracteres = len(mat_item)
						'Posição do separador
							separador = Instr(mat_item,";")
						'Quantidade de materiais
							mat_item_qtd = mid(mat_item,separador+1,mat_caracteres)
						'Código do material
							mat_item = mid(mat_item,1,separador-1)
						
						
						'mat_item_qtd = replace(right(mat_item,(separador)),";","")
						'mat_item = replace(left(mat_item,(separador)),";","")
					
				if IsNumeric(mat_item) then
					Response.Write("<br>" & mat_item)
					xsql ="SELECT * FROM TBL_protocolo_material WHERE A061_codmat ='"&cd_protocolo&"' AND A061_codmat ='"&mat_item&"'"
					Set rs_mat = dbconn.execute(xsql)
					
						
						'if not rs_mat.EOF then
							'Altera o Material
						'	cd_prot_mat = rs_mat("cd_codigo")
						'	xsql ="up_protocolo_material @cd_prot_mat = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @material='"&mat_item&"', @primeira=N"
						'	dbconn.execute(xsql)
						'	'response.write(" "&numero&" existe")
						'	mensagem3 = "Material já existe."
						'else
						if rs_mat.EOF Then
							'Insere o Material
							xsql_mat ="up_protocolo_material_insert @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"',@material='"&mat_item&"', @quantidade='"&mat_item_qtd&"'"
							dbconn.execute(xsql_mat)
							'response.write(" "&numero&" não existe")
							mensagem4 = "material inserido."
						End if
				end if
					numero = numero + 1
				Next
			end if
		
		'*** Patrimonios *** 
		'++++++++++++++++++++++++++
		'+++ INSERE PATRIMONIOS +++
		'++++++++++++++++++++++++++
		numero = 1
		pat_array = split(subtotal_patrimonio,",")
		For Each pat_item In pat_array
			
		if IsNumeric(pat_item) then
			xsql ="SELECT * FROM TBL_protocolo_patrimonio WHERE cd_protocolo ='"&cd_protocolo&"' AND cd_unidade = '"&cd_unidade&"' AND cd_patrimonio ='"&pat_item&"'"
			Set rs_pat = dbconn.execute(xsql)
				
				if not rs_pat.EOF then
					'Altera o Procedimento
					cd_prot_pat = rs_pat("cd_codigo")
					xsql ="up_protocolo_patrimonio_update @cd_prot_pat = '"&int(cd_prot_pat)&"', @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @patrimonio='"&pat_item&"'"
					dbconn.execute(xsql)
					''response.write(" "&numero&" existe")
					mensagem1 = "Patrimonio já existe."
				elseif rs_pat.EOF Then
					'Insere o procedimento
					xsql_pat ="up_protocolo_patrimonio_insert @protocolo='"&cd_protocolo&"', @unidade='"&cd_unidade&"', @patrimonio='"&pat_item&"'"
					dbconn.execute(xsql_pat)
					'response.write(" "&numero&" não existe")
					mensagem2 = "Patrimonio inserido."				
				End if
		end if
				
			numero = numero + 1
		Next
	response.redirect "../../protocolo.asp?tipo=ver&cd_unidade="&zero(cd_unidade)&"&cd_protocolo="&int(cd_protocolo)&"&cd_digito="&cd_digito&"&msg="&mensagem&""

	Elseif acao = "excluir" Then
'***************************************************************************************************************************************************************
'*** EXCLUIR PROTOCOLO ****************************************************************************************************************************************
'*************************************************************************************************************************************************************
	
		if obj_excl = 1 Then
			' Exclui o protocolo
		
		elseif obj_excl = 2 Then	
			'xsql ="up_protocolo_delete @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"'"
			'dbconn.execute(xsql)
		
		elseif obj_excl = 3 Then
			'Exclui o material
			xsql ="up_protocolo_material_delete @protocolo="&cd_protocolo&", @unidade="&cd_unidade&",@cd_codigo="&cd_material&""
			dbconn.execute(xsql)
			response.redirect("../../protocolo.asp?tipo=digitacao&cd_protocolo="&cd_protocolo&"&cd_unidade="&zero(cd_unidade)&"&cd_digito="&cd_digito&"&mensagem="&mensagem&"")
		'end if
			
			
	'****************************
	'*** EXCLUIR PROCEDIMENTO ***
	'****************************
		elseif obj_excl = 6 Then
			'Exclui o procedimento
			xsql ="up_protocolo_procedimento_delete @cd_codigo="&cd_codigo&""
			dbconn.execute(xsql)
			response.redirect("../../protocolo.asp?tipo=digitacao&cd_protocolo="&cd_protocolo&"&cd_unidade="&zero(cd_unidade)&"&cd_digito="&cd_digito&"&mensagem="&mensagem&"")
		
		elseif obj_excl = 7 Then
		'	'Exclui o procedimento
			xsql ="up_protocolo_patrimonio_delete @cd_codigo="&cd_codigo&""
			dbconn.execute(xsql)
			response.redirect("../../protocolo.asp?tipo=digitacao&cd_protocolo="&cd_protocolo&"&cd_unidade="&zero(cd_unidade)&"&cd_digito="&cd_digito&"&mensagem="&mensagem&"")
		end if
	
	'DELETE FROM TBL_Protocolo_Procedimento
	'WHERE     (A002_numpro = 3185) AND (a053_codung = 15)
	
	
	
	
	'response.write("OK, alterado")
	'mensagem = "Ok, protocolo alterado com sucesso"
	
	'Elseif acao = "excluir" Then
	'	xsql ="up_protocolo_delete @A002_numseq='"&cd_codigo&"', @cd_procedimento='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm_1&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"'"
	'	dbconn.execute(xsql)
'	
	'	DELETE FROM TBL_Protocolo_Procedimento
	'	WHERE     (A002_numseq = 17226) AND (A058_codpro = 45020060)
	
	
	Else
	
	response.write("Erro, não passou")
	End if
'Else
	if cd_user <> 460 then
	response.redirect("../../protocolo.asp?tipo=digitacao&mensagem="&mensagem&"")
	end if
'end if
end if 'evita que os dados sejam gravados no Banco de dados, fins de desenvolvimento%>
<br>
<br>
acao: <%=acao%><br>
modo: <%=modo%><br>
tipo: <%=tipo%><br>
<br>
Unidade: <%=cd_unidade%><br>
protocolo: <%=cd_protocolo%>-<%=cd_digito%><br><br>

código: <%=cd_codigo%><br>
Nome: <%=nm_nome%><br>
Idade: <%=cd_idade%><br>
registro: <%=nm_registro%><br>
cd_convenio: <%=cd_convenio%><br>
sexo: = <%=nm_sexo%><br>
cirug_realizada?: <%=nm_cirug_realizada%><br>
dt_cirurgia: <%=dt_cirurgia%><br>
nm_agenda_cirurgia: <%=nm_agenda_cirurgia%><br>
hr_agenda: <%=dt_hr_agenda%><br>
hora_inicio: <%=dt_hr_inicio%><br>
hora_fim: <%=dt_hr_fim%><br>
CRM: <%=cd_crm_1%><br>
cd_rack: <%=cd_rack%><br>
Especialidade: <%=cd_especialidade%><br>
procedimentos: <%=procedimentos_total%><br>
Materiais: <%=materiais_total%><br>
Cod material = <%=cd_material%><br>
Objeto de exclusao = <%=obj_excl%><br>
Patrimonio: <%=subtotal_patrimonio%><br>
CO: <%=cd_co%>
///////<br>
M1:<%=mensagem1%><br>
M2:<%=mensagem2%><br>
M3:<%=mensagem3%><br>
M4:<%=mensagem4%><br>
