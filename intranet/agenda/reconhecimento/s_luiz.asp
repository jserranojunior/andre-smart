<%'informacao_bruta = Replace(request("informacao_bruta"),vbcrlf,";<br>")
'informacao_bruta = request("informacao_bruta")
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
'session("data_selecionada") = dia_sel&"/"&mes_sel&"/"&ano_sel
if dia_sel = "" AND mes_sel = "" and ano_sel = "" then
	dia_sel = day(now)
	mes_sel = month(now)
	ano_sel = year(now)
end if

session("data_selecionada") = ""'dt_data

if informacao_bruta <> "" Then
	'*** Encontra a posição do dado a inserir na tabela ****************************************
	pos_hsli = instr(1,informacao_bruta,"Unidade Itaim",1)
	pos_hslm = instr(1,informacao_bruta,"Unidade Morumbi",1)
	pos_hsla = instr(1,informacao_bruta,"Unidade Analia",1)
	
	pos_reserva = instr(1,informacao_bruta,"RESERVA       :",1)
	pos_paciente = instr(pos_reserva,informacao_bruta,"PACIENTE      :",1)
	pos_convenio = instr(pos_paciente,informacao_bruta,"CONVENIO      :",1)
	pos_situacao = instr(pos_convenio,informacao_bruta,"SITUACAO      :",1)
	pos_datareserv = instr(pos_situacao,informacao_bruta,"DATA RESERVA  :",1)
	pos_dataintern = instr(pos_datareserv,informacao_bruta,"DATA INTERNAC.:",1)
	pos_datacirurg = instr(pos_dataintern,informacao_bruta,"DATA CIRURGIA :",1)
	pos_cirurgia = instr(pos_datacirurg,informacao_bruta,"CIRURGIA      :",1)
	pos_medico = instr(pos_cirurgia,informacao_bruta,"MEDICO        :",1)
	pos_medcredenc = instr(pos_medico,informacao_bruta,"MEDICO CREDENCIADO NO CONVENIO?",1)
	pos_patologista = instr(pos_medcredenc,informacao_bruta,"PATOLOGISTA   :",1)
	pos_proteses = instr(pos_patologista,informacao_bruta,"PROTESES      :",1)
	pos_centro = instr(pos_proteses,informacao_bruta,"CENTRO        :",1)
	pos_anestesia = instr(pos_centro,informacao_bruta,"Anestesia",1)
	
	pos_identificador1 = instr(pos_anestesia,informacao_bruta,"CIRURGIAS :",1)
	pos_amb = instr(pos_anestesia,informacao_bruta,"AMB :",1)
	
	pos_identificador2 = instr(pos_amb,informacao_bruta,"MATERIAIS CONSIGNADOS / MEDICAMENTOS :",1)
		pos_material = instr(pos_amb,informacao_bruta,"Material  :",1)
		if pos_material > 0 then 
			pos_fornecedor = instr(pos_material,informacao_bruta,"Fornecedor:",1)
			pos_quantidade = instr(pos_fornecedor,informacao_bruta,"Quantidade:",1)
		else
			pos_quantidade = pos_amb
		end if
	
	pos_identificador3 = instr(pos_identificador2,informacao_bruta,"LISTA DE EQUIPAMENTOS :",1)
		if pos_material > 0 then 
			pos_equip = instr(pos_quantidade,informacao_bruta,"Equip.:",1)
			pos_caixa = instr(pos_equip,informacao_bruta,"Caixa.:",1)
		else
			pos_caixa = pos_quantidade
		end if
	pos_obs = instr(pos_caixa,informacao_bruta,"OBSERVACOES DA RESERVA :",1)
	pos_atencao = instr(pos_obs,informacao_bruta,"ATENÇÃO:",1)
	
	
	'*** tamanho do intervalo a excluir ********************************************************
	x_informacao_bruta = len(informacao_bruta)
	x_reserv = len("RESERVA       :")
	x_paciente = len("PACIENTE       :")
	x_convenio = len("CONVENIO      :")
	x_situacao = len("SITUACAO      :")
	x_datareserv = len("DATA RESERVA  :")
	x_dataintern = len("DATA INTERNAC.:")
	x_datacirurg = len("DATA CIRURGIA :")
	x_cirurgia = len("CIRURGIA      :")
	x_medico = len("MEDICO        :")
	x_medcredenc = len("MEDICO CREDENCIADO NO CONVENIO?")
	x_patologista = len("PATOLOGISTA   :")
	x_proteses = len("PROTESES      :")
	x_centro = len("CENTRO        :")
	x_anestesia = len("Anestesia")
	
	x_amb = len("AMB :")
	'x_amb = len("CENTRO        :")
	
	x_material = len("Material  : ")
	'x_fornecedor = len("Fornecedor:")
	'x_quantidade = len("Quantidade:")
	
	x_equip = len("Equip.:")
	x_caixa = len("Caixa.:")
	x_obs = len("OBSERVACOES DA RESERVA :")
	x_atencao = len("ATENÇÃO:")
	
	x_identificador1 = len("CIRURGIAS :")
	x_identificador2 = len("MATERIAIS CONSIGNADOS / MEDICAMENTOS :")
	x_identificador3 =  len("LISTA DE EQUIPAMENTOS :")
	
	'*** define o intervalo da string que será mostrado ****************************************
	intervalo_reserva = pos_paciente - pos_reserva
	intervalo_paciente = pos_convenio - pos_paciente
	intervalo_convenio = pos_situacao - pos_convenio
	intervalo_situacao = pos_datareserv - pos_situacao
	intervalo_datareserv = pos_dataintern - pos_datareserv
	intervalo_dataintern = pos_datacirurg - pos_dataintern
	intervalo_datacirurg = pos_cirurgia - pos_datacirurg
	intervalo_cirurgia = pos_medico - pos_cirurgia
	intervalo_medico = pos_medcredenc - pos_medico
	intervalo_medcredenc = pos_patologista - pos_medcredenc
	intervalo_patologista = pos_proteses - pos_patologista
	intervalo_proteses = pos_centro - pos_proteses
	intervalo_centro = pos_anestesia - pos_centro
	intervalo_anestesia = pos_identificador1 - pos_anestesia
	intervalo_amb = pos_identificador2 - pos_amb
	
	'intervalo_equipamentos = pos_
	intervalo_materiais = pos_identificador3 - pos_material
		intervalo_material = pos_fornecedor - pos_material
		intervalo_fornecedor = pos_quantidade - pos_fornecedor
		intervalo_quantidade = pos_equip - pos_quantidade
	intervalo_equipamentos = pos_obs - pos_identificador3
		intervalo_equip = pos_caixa - pos_equip
		intervalo_caixa = pos_obs - pos_caixa
	intervalo_obs = pos_atencao - pos_obs
	
	
	'*** Mostra a string final *****************************************************************
	str_reserva = mid(informacao_bruta,(pos_reserva+x_reserv),(intervalo_reserva-x_reserv))
	str_paciente = mid(informacao_bruta,(pos_paciente+x_paciente),(intervalo_paciente-x_paciente))
	str_convenio = mid(informacao_bruta,(pos_convenio+x_convenio),(intervalo_convenio-x_convenio))
	str_situacao = mid(informacao_bruta,(pos_situacao+x_situacao),(intervalo_situacao-x_situacao))
	str_datareserv = mid(informacao_bruta,(pos_datareserv+x_datareserv),(intervalo_datareserv-x_datareserv))
	str_dataintern = mid(informacao_bruta,(pos_dataintern+x_dataintern),(intervalo_dataintern-x_dataintern))
	str_datacirurg = mid(informacao_bruta,(pos_datacirurg+x_datacirurg),(intervalo_datacirurg-x_datacirurg))
		pos_hifen = instr(1,str_datacirurg,"-",1)
			str_horacirurg = replace(mid(str_datacirurg,pos_hifen+1,len(str_datacirurg)),",",":")
			str_datacirurg = mid(str_datacirurg,1,pos_hifen-1)&str_horacirurg
	
	str_cirurgia = mid(informacao_bruta,(pos_cirurgia+x_cirurgia),(intervalo_cirurgia-x_cirurgia))
	str_medico = mid(informacao_bruta,(pos_medico+x_medico),(intervalo_medico-x_medico))
	str_medcredenc = mid(informacao_bruta,(pos_medcredenc+x_medcredenc),(intervalo_medcredenc-x_medcredenc))
	str_patologista = mid(informacao_bruta,(pos_patologista+x_patologista),(intervalo_patologista-x_patologista))
	str_proteses = mid(informacao_bruta,(pos_proteses+x_proteses),(intervalo_proteses-x_proteses))
	str_centro = mid(informacao_bruta,(pos_centro+x_centro),(intervalo_centro-x_centro))
	'str_centro = mid(informacao_bruta,(pos_centro),(intervalo_centro))
	str_amb = mid(informacao_bruta,(pos_amb+x_amb),(intervalo_amb-x_amb))
		'for each c as char in str_amb 
		'  if isnumeric(c) then
		'     s1+=c.tostring
		'  end if
		'next

	
	'str_materiais = mid(informacao_bruta,pos_material+x_material,pos_identificador3-pos_material)
	str_materiais = mid(informacao_bruta,pos_material+x_material,intervalo_materiais-x_material)
		if pos_material > 0 then
		mat_verificar = instr(1,str_materiais,";",1)
			while mat_verificar > 0
				pontos = pontos&","&mat_verificar
				str_materiais = replace(str_materiais,";",VbCrLf)
				mat_verificar = instr(1,mid(valor,mat_verificar,len(mat_verificar)),";",1)
			wend
		else
			str_materiais = ""
		end if
		
		'str_material = mid(informacao_bruta,(pos_material+x_material),(intervalo_material-x_material))
		'str_fornecedor = mid(informacao_bruta,(pos_fornecedor+x_fornecedor),(intervalo_fornecedor-x_fornecedor))
		'str_quantidade = mid(informacao_bruta,(pos_quantidade+x_quantidade),(intervalo_quantidade-x_quantidade))
	
	str_equipamentos = mid(informacao_bruta, pos_identificador3+x_identificador3, intervalo_equipamentos-x_identificador3)
		'equip_verificar = instr(1,str_equipamentos,";",1)
			'while equip_verificar > 0
			'	pontos = pontos&","&equip_verificar
				'eq_salto_verificar = instr(1,str_equipamentos,VbCrLf,1) '// Verifica salto de linha no início//
				'if 
				'str_equipamentos = replace(str_equipamentos,VbCrLf,"...")
				'str_equipamentos = replace(str_equipamentos,";",VbCrLf)
				'equip_verificar = instr(1,mid(valor,equip_verificar,len(equip_verificar)),";",1)
			'wend
		
		'str_equip = mid(informacao_bruta,(pos_equip+x_equip),(intervalo_equip-x_equip))
		'str_caixa = mid(informacao_bruta,(pos_caixa+x_caixa),(intervalo_caixa-x_caixa))
		
	str_obs = mid(informacao_bruta,(pos_obs+x_obs),(intervalo_obs-x_obs))
	'*******************************************************************************************
end if
%>
