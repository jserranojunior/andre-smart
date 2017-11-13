<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
   <!--include file="../includes/inc_area_restrita.asp"-->
<%
'***** Variáveis ******

data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")

acao = request("acao") 
cd_unidade = request("cd_unidade")
cd_protocolo = request("cd_protocolo")
cd_digito = request("cd_digito")
cd_codigo = request("cd_codigo")



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

	dt_hora_agenda = request("dt_hora_agenda")
	dt_minuto_agenda = request("dt_minuto_agenda")
dt_hr_agenda = dt_hora_agenda&":"&dt_minuto_agenda&":00"

	dt_hora_inicio = request("dt_hora_inicio")
	dt_minuto_inicio = request("dt_minuto_inicio")
dt_hr_inicio = dt_hora_inicio&":"&dt_minuto_inicio

	dt_hora_fim = request("dt_hora_fim")
	dt_minuto_fim = request("dt_minuto_fim")
dt_hr_fim = dt_hora_fim&":"&dt_minuto_fim

'cd_crm = request("cd_crm")
	cd_crm_1 = request("cd_crm_1")

cd_rack = request("cd_rack")
	cd_rack_1 = request("cd_rack_1")

cd_especialidade = request("cd_especialidade")
	cd_especialidade_1 = request("cd_especialidade_1")

cd_procedimento = request("cd_procedimento")
	cd_procedimento_1 = request("cd_procedimento_1")

procedimentos = request("procedimentos")
	if right(procedimentos,1) = "," Then
	proc_len = len(procedimentos)
	proc_len = proc_len - 1
	
		procedimentos = Left(procedimentos,proc_len)
	end if
	
materiais = request("materiais")
	if right(materiais,1) = "," Then
	mat_len = len(materiais)
	mat_len = mat_len - 1
	
		materiais = Left(materiais,mat_len)
	end if


'************ Condições ************************
if cd_convenio_1 <> "" Then
	cd_convenio = cd_convenio_1
end if

if cd_crm_1 <> "" Then
	strsql ="Select * from TBL_medicos where a055_crmmed="&cd_crm_1&""
		Set rs_med = dbconn.execute(strsql)
		if not rs_med.EOF  then
		cd_crm = rs_med("A055_codmed")
		end if
end if	
'if cd_crm_1 <> "" Then
'	cd_crm = cd_crm_1
'end if

if cd_rack_1 <> "" Then
	cd_rack = cd_rack_1
end if
	if cd_rack = "" Then
	cd_rack = "11"
	end if

if cd_especialidade_1 <> "" Then
	cd_especialidade = cd_especialidade_1
end if

'if cd_procedimento_1 <> "" Then
'	cd_procedimento = cd_procedimento_1
'end if 

if dt_hr_agenda = ":" Then
	dt_hr_agenda = NULL
End if

if dt_hr_inicio = ":" Then
	dt_hr_inicio = NULL
End if

if dt_hr_fim = ":" Then
	dt_hr_fim = NULL
End if

	'***** Ações ******
	If acao = "inserir" Then

	xsql ="up_protocolo_insert @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm_1&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"'"',@procedimentos='"&procedimentos&"'"
	dbconn.execute(xsql)
	response.write("OK, inserido")
	mensagem = "OK, protocolo inserido com sucesso"
'	response.redirect "../protocolo.asp"
		'Verifica os procedimentos
		
		'*** Procedimentos ***
		numero = 1
		proc_array = split(procedimentos,",")
		For Each proc_item In proc_array
			
			if IsNumeric(proc_item) then
				Response.Write("<br>" & proc_item)
				xsql ="SELECT * FROM TBL_protocolo_procedimento WHERE A002_numseq ='"&cd_protocolo&"' AND A058_codpro ='"&proc_item&"'"
				Set rs_proc = dbconn.execute(xsql)
					
					if not rs_proc.EOF then
						'Altera o Procedimento
						cd_prot_proc = rs_proc("cd_codigo")
						xsql ="up_protocolo_procedimento_update @cd_prot_proc = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @procedimento='"&proc_item&"', @primeira=N"
						dbconn.execute(xsql)
						'response.write(" "&numero&" existe")
						mensagem = "Procedimento já existe."
					elseif rs_proc.EOF Then
						'Insere o procedimento
						
							if numero = 1 then 
							proc_cat = "S"
							else
							proc_cat = "N"
							End if
						
						xsql_proc ="up_protocolo_procedimento_insert @protocolo='"&cd_protocolo&"', @procedimento='"&proc_item&"', @primeira='"&proc_cat&"'"
						dbconn.execute(xsql_proc)
						'response.write(" "&numero&" não existe")
						mensagem = "Procedimento inserido."
					End if
					
				numero = numero + 1
			'else
			'		response.write(mat_item)
			end if
		Next
	
	
		'*** Materiais *** 
		if materiais <> "" then
			mat_array = split(materiais,",")
			For Each mat_item In mat_array
				
				separador = Instr(mat_item,";")
				mat_item_qtd = replace(right(mat_item,(separador-1)),";","")
				mat_item = left(mat_item,(separador-1))
				
				if IsNumeric(mat_item) then
					xsql ="SELECT * FROM TBL_protocolo_material WHERE A002_numseq ='"&cd_protocolo&"' AND A061_codmat ='"&mat_item&"'"
					Set rs_mat = dbconn.execute(xsql)
						
						if rs_mat.EOF then 'Verifica se haverá duplicação
							xsql_mat ="up_protocolo_material_insert @protocolo='"&cd_protocolo&"', @material='"&mat_item&"', @quantidade='"&mat_item_qtd&"'"
							dbconn.execute(xsql_mat)
							mensagem2 = "material inserido."
						end if
				else
					response.write(mat_item)
				end if
				
			Next
		end if

	Elseif acao = "editar" Then
	xsql ="up_protocolo_update @cd_codigo='"&cd_codigo&"', @cd_protocolo='"&int(cd_protocolo)&"', @cd_unidade='"&cd_unidade&"',@nm_nome='"&nm_nome&"', @cd_idade='"&cd_idade&"', @nm_registro='"&nm_registro&"',@cd_convenio='"&cd_convenio&"',@nm_sexo='"&nm_sexo&"', @nm_cirug_realizada='"&nm_cirug_realizada&"',@dt_cirurgia='"&dt_cirurgia&"', @nm_agenda_cirurgia='"&nm_agenda_cirurgia&"',@dt_hr_agenda='"&dt_hr_agenda&"',@dt_hr_inicio='"&dt_hr_inicio&"',@dt_hr_fim='"&dt_hr_fim&"',@cd_crm='"&cd_crm_1&"',@cd_rack='"&cd_rack&"',@cd_especialidade='"&cd_especialidade&"'"
	dbconn.execute(xsql)
	response.write("OK, alterado")
	mensagem = "Ok, protocolo alterado com sucesso"
	
		'Verifica os materiais
		'*** Procedimentos ***
		numero = 1
		proc_array = split(procedimentos,",")
		For Each proc_item In proc_array
			
			Response.Write("<br>" & proc_item)
			xsql ="SELECT * FROM TBL_protocolo_procedimento WHERE A002_numseq ='"&cd_protocolo&"' AND A058_codpro ='"&proc_item&"'"
			Set rs_proc = dbconn.execute(xsql)
				
				if not rs_proc.EOF then
					'Altera o Procedimento
					cd_prot_proc = rs_proc("cd_codigo")
					xsql ="up_protocolo_procedimento_update @cd_prot_proc = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @procedimento='"&proc_item&"', @primeira=N"
					dbconn.execute(xsql)
					'response.write(" "&numero&" existe")
					mensagem = "Procedimento já existe."
				elseif rs_proc.EOF Then
					'Insere o procedimento
					
						if numero = 1 then 
						proc_cat = "S"
						else
						proc_cat = "N"
						End if
					
					xsql_proc ="up_protocolo_procedimento_insert @protocolo='"&cd_protocolo&"', @procedimento='"&proc_item&"', @primeira='"&proc_cat&"'"
					dbconn.execute(xsql_proc)
					'response.write(" "&numero&" não existe")
					mensagem = "Procedimento inserido."
				End if
				
			numero = numero + 1
		Next
	
	
		'*** Materiais *** 
		if materiais <> "" then
		'numero = 1
			mat_array = split(materiais,",")
			For Each mat_item In mat_array
				
				
				separador = Instr(mat_item,";")
				mat_item_qtd = replace(right(mat_item,(separador-1)),";","")
				mat_item = left(mat_item,(separador-1))
				
				
				if IsNumeric(mat_item) then
					Response.Write("<br>" & mat_item)
					xsql ="SELECT * FROM TBL_protocolo_material WHERE A002_numseq ='"&cd_protocolo&"' AND A061_codmat ='"&mat_item&"'"
					Set rs_mat = dbconn.execute(xsql)
						
						if not rs_mat.EOF then
							'Altera o Material
							cd_prot_mat = rs_mat("cd_codigo")
							xsql ="up_protocolo_material @cd_prot_mat = '"&cd_prot_proc&"', @protocolo='"&cd_protocolo&"', @material='"&mat_item&"', @primeira=N"
							dbconn.execute(xsql)
							'response.write(" "&numero&" existe")
							'mensagem2 = "Material já existe."
						elseif rs_mat.EOF Then
							'Insere o Material
							xsql_mat ="up_protocolo_material_insert @protocolo='"&cd_protocolo&"', @material='"&mat_item&"', @quantidade='"&mat_item_qtd&"'"
							dbconn.execute(xsql_mat)
							'response.write(" "&numero&" não existe")
							'mensagem2 = "material inserido."
						End if
						
					'numero = numero + 1
				else
					response.write(mat_item)
				end if
				
			Next
		end if
		
		
		
	'response.redirect "../protocolo.asp?tipo=ver&cd_unidade="&zero(cd_unidade)&"&cd_protocolo="&int(cd_protocolo)&"&cd_digito="&cd_digito&"&msg="&mensagem&""

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
	'response.redirect("../protocolo.asp?tipo=digitacao")
'end if
%>
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
CRM: <%=cd_crm%><br>
cd_rack: <%=cd_rack%><br>
Especialidade: <%=cd_especialidade%><br>
procedimentos: <%=procedimentos%><br>
Materiais: <%=materiais%><br><br>
Mensagem: <%=mensagem%><br>
Mensagem2: <%=mensagem2%>
