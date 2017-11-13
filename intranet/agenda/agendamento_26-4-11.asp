<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!-- TESTE DE LAYOUT - contagem de dados para criação da tabela e posicionamento dos agendamentos -->
<html>
<head>


<title>Agenda</title>
<script type="text/javascript"> 
	function pasteClipboard() { 
			txt = document.form.informacao_bruta.createTextRange();
			txt.execCommand('Paste');
		} 
		
	function mostra_campos() {
		document.getElementById('nome').style.zIndex="1";
		}
		
	function esconde_campos() {
		document.getElementById('nome').style.zIndex="-1";
		}
		
	function agendamento_abre(situacao) {
		document.getElementById('campo_agendamento').style.display=situacao;
		}
	
	function mostra_agendamento(situacao,topo,esquerda,campo) {		
		document.getElementById('mostra_agendamento').style.display=situacao;
		document.getElementById('mostra_agendamento').style.top=topo
		document.getElementById('mostra_agendamento').style.left=esquerda;
		setTimeout('document.getElementById("mostra_agendamento").style.display="none"',15000);
		
			document.Form.ajax.value=campo;{	
			 var oHTTPRequest_agendada = createXMLHTTP(); 
		     oHTTPRequest_agendada.open("post", "agenda/ajax/agendados.asp", true);
		     oHTTPRequest_agendada.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		     oHTTPRequest_agendada.onreadystatechange=function(){
		      if (oHTTPRequest_agendada.readyState==4){
		         document.all.divAgendamento.innerHTML = oHTTPRequest_agendada.responseText;}}
		       oHTTPRequest_agendada.send("ajax=" + Form.ajax.value);
		 		}
				}
		
	function esconde_agendamento() {		
		document.getElementById('ajax').style.display='none';
	}

	
	
	
		
	      
	function mostra(event)         {    
		var diferenca;        
			document.getElementById("flutua").style.display = "inline"        
			diferenca = document.getElementById("flutua").offsetHeight + 0;       
			document.getElementById("flutua").style.top = event.offsetY - diferenca;    
			document.getElementById("flutua").style.left = event.offsetX;
			document.getElementById("flutuador").flutua_pos.value=event.offsetX;
	}
			    
		function some()         {         
			document.getElementById("flutua").style.display = "none";
		} 
</script>





<%dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
	if dia_sel="" OR mes_sel="" OR  ano_sel= "" Then
		dia_sel = day(now)
		mes_sel = month(now)
		ano_sel = year(now)
	end if 
informacao_bruta = request("informacao_bruta")%>
</head>
<!--#include file="../js/ajax.js"-->
<!--#include file="../agenda/ajax/agenda.js"-->
<body >

<table cellspacing="2" cellpadding="2" border="0" align="center" class="no_print">
<form action="../agenda.asp" name="form" id="form" method="post">
<input type="hidden" name="tipo" value="agendamento">
	<tr>
		<td><textarea cols="0" rows="1" name="informacao_bruta" onclick="document.formsubmit;"></textarea>
		<input type="submit" name="Colar" value="Importar" onclick="pasteClipboard()"></td>
	</tr>
	<tr>
		<td></td>
	</tr>
</form>
	<tr>
		<td><%'=informacao_bruta%></td>
	</tr>
	<tr>
		<td align="center">
			<%'str_analia = instr(informacao_bruta,"Unidade Analia")
			'str_itaim = instr(informacao_bruta,"Unidade Itaim")
			'str_morumbi = instr(informacao_bruta,"Unidade Morumbi")
			
			str_sluiz = instr(informacao_bruta,"INFORMACAO SOBRE RESERVA")
				if sluiz > 0 then
					'**** 
					'*** DEFINE A POSIÇÃO DO NÚMERO DA RESERVA NO TEXTO ***
					reserva = instr(1,informacao_bruta,"RESERVA       :",1)
					paciente = instr(reserva,informacao_bruta,"PACIENTE      :",1)
					'*** Exclui a palavra: RESERVA ***
					x_reserv = len("RESERVA       :")
					'*** Intervalo da reserva, encontra o número da reserva ***
					intv_reserva = paciente - reserva
					'*** ***
					nr_reserva = mid(informacao_bruta,(reserva+x_reserv),(intv_reserva-x_reserv))
					
					
				end if
			str_leforte = instr(informacao_bruta,"Hosp. Leforte")
			str_villa = instr(informacao_bruta,"RINT51150A") 'OK
			
			if str_sluiz > 0 then%>
				<!--#include file="../agenda/reconhecimento/s_luiz.asp"-->
			<%elseif str_analia > 0 Then
				response.write("Hospital S.Luiz Anália")%>
				<!--#include file="../agenda/reconhecimento/s_luiz.asp"-->
			<%elseif str_itaim > 0 Then
				response.write("Hospital S.Luiz Itaim")%>
				<!--#include file="../agenda/reconhecimento/s_luiz.asp"-->
			<%elseif str_morumbi > 0 then
				response.write("Hospital S.Luiz Morumbi")%>
				<!--#include file="../agenda/reconhecimento/s_luiz.asp"-->
			<%elseif str_leforte > 0 then
				response.write("Hospital Leforte")%>
				
			<%elseif str_villa > 0 then
				response.write("Hospital Villa Lobos")%>
				
			<%else
				response.write("Não foi possível determinar qual a unidade.<br>Por favor, preencha a formulário abaixo.")
			end if%>
		</td>
	</tr>
	<!--tr>
		<td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>
	</tr>
	<tr>
		<td>
			<%'response.write(pos_amb)%>			
		</td>
	</tr-->
	<!--tr>
		<td>Villa: <%=str_villa%><br>
		S. Luiz Itaim: <%=str_itaim%><br>
		S. Luiz Morumbi: <%=str_morumbi%><br>
		S. Luiz Anália: <%=str_analia%><br>
		Leforte: <%=str_leforte%><br>
		</td>
	</tr-->
</table>
<br>


<%'=informacao_bruta%>

		
			<!-- onClick="controlaCamada('nome')"-->
			
			<table border="1" style="border-collapse=collapse;">
				<form action="agendamento.asp" name="Form" id="Form">
				<input type="text" name="ajax">
				<tr>
					<td colspan="100%" align="center">AGENDA</td>
				</tr>
				<tr>
					<td colspan="100%" align="center">Hoje é: <%=zero(dia_sel)%> de <%=mes_selecionado(mes_sel)%> de <%=ano_sel%></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				<%'*********************************************************
				'*** ! - Conta quantos hospitais são clientes da empresa ***
				'***********************************************************
				
				'*** A - Conta as unidades ativas ***
				'xsql = "SELECT COUNT(nm_unidade) AS qtd_unidades_ativas FROM TBL_unidades WHERE (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011')"
				xsql = "SELECT COUNT(nm_unidade) AS qtd_unidades_ativas FROM VIEW_agenda_unidade_ordem WHERE (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011') GROUP BY cd_ordem,nm_unidade ORDER BY cd_ordem,nm_unidade"
				
				Set rs_und = dbconn.execute(xsql)
				while not rs_und.EOF
					qtd_unidades_ativas = rs_und("qtd_unidades_ativas")
				rs_und.movenext
				wend
					
					'*** B - Conta os racks de cada unidade ativa***
					xsql = "SELECT * FROM TBL_unidades WHERE (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011')"
					'xsql = "SELECT * FROM VIEW_agenda_unidade_ordem WHERE (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011') GROUP BY cd_ordem,nm_unidade,nm_sigla,dt_data,cd_unidade,cd_codigo,cd_hospital,dt_inicio_contrato,dt_fim_contrato ORDER BY cd_ordem,nm_unidade"
					Set rs_und = dbconn.execute(xsql)
					while not rs_und.EOF
						cd_unidade_ativa = rs_und("cd_codigo")
						
							'*** C - Conta os racks das unidades ativas
							xsql = "SELECT * FROM TBL_unidades_racks WHERE (cd_unidade = "&cd_unidade_ativa&") AND (dt_inicio_uso >= '1/1/2000') AND (dt_fim_uso IS NULL) OR (cd_unidade = "&cd_unidade_ativa&") AND (dt_inicio_uso >= '1/1/2000') AND (dt_fim_uso <= '12/31/2011')"
							Set rs_rack = dbconn.execute(xsql)
							while not rs_rack.EOF
								qtd_racks = qtd_racks + int(1)
								total_racks = total_racks + 1
								conjunto_racks = conjunto_racks &":"&rs_rack("cd_rack")
							rs_rack.movenext
							wend
								qtd_racks_ativos = qtd_racks_ativos + qtd_racks
								conjunto_racks = mid(conjunto_racks,2,len(conjunto_racks))
								conjunto_total = conjunto_total &";"&cd_unidade_ativa&","&conjunto_racks&","&qtd_racks
						qtd_racks = "0"
						conjunto_racks = ""
					rs_und.movenext
					wend%>
				<%conjunto_total = mid(conjunto_total,2,len(conjunto_total))%>
				</tr>
				
				<!--tr><td colspan="8">***<%=conjunto_total%></td></tr-->
				<tr>
					<td>&nbsp;</td><!-- *** UNIDADES *** -->
					<%'*** Separa e mostra cada unidade ***
					unidades_array = split(conjunto_total,";")
					For Each unidades_item_junto In unidades_array
						unidades_item_junto_len = len(unidades_item_junto)
						
						separador1 = instr(unidades_item_junto,",")
						separador2 = instr(mid(unidades_item_junto,(separador1)+1,unidades_item_junto_len),",")
						
						strunidade = mid(unidades_item_junto,1,separador1-1)
						strracks = mid(unidades_item_junto,separador1+1,separador2-1)
						strqtd_racks = mid(unidades_item_junto,(separador1+separador2)+1,separador2-1)
							'*** Nome da Unidade ***
							xsql = "SELECT * FROM TBL_unidades WHERE cd_codigo = "&strunidade&""
							Set rs_und = dbconn.execute(xsql)
								if not rs_und.EOF Then
									if strqtd_racks > 2 then
										nm_unidade = rs_und("nm_unidade")
									else
										nm_unidade = rs_und("nm_sigla")
									end if
								end if
									
						response.write("<td colspan='"&strqtd_racks&"' align='center'>"&nm_unidade&"</td>")
						
						'qtd_colunas = qtd_rack + int(strrack)
						
					next%>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td><!-- *** RACKS *** -->
					<%'*** Separa e mostra cada unidade ***
						unidades_array = split(conjunto_total,";")
						For Each unidades_item_junto In unidades_array
							unidades_item_junto_len = len(unidades_item_junto)
							
							separador1 = instr(unidades_item_junto,",")
							separador2 = instr(mid(unidades_item_junto,(separador1)+1,unidades_item_junto_len),",")
							
							'strunidade = mid(unidades_item_junto,1,separador1-1)
							strracks = mid(unidades_item_junto,separador1+1,separador2-1)
							strqtd_racks = mid(unidades_item_junto,(separador1+separador2)+1,separador2-1)
							
								racks_array = split(strracks,":")
								For Each  racks_item_junto In racks_array
									racks_item_junto_len = len(racks_item_junto)
									response.write("<td align='center'>"&racks_item_junto&"</td>")
								'response.write("<td>"&strunidade&" *** "&strracks&" *** "&strqtd_racks&"<br>"&separador1&"-"&separador2&"-"&separador3&"</td>")
								next
						next%>
					<td>&nbsp;</td>
				</tr>
				<%topo = 150
				nome_campo = 0
				for horarios = 1 to 24 '*** linhas (24horas) ***
					if horarios = 19 then
						hora_ag =  0
					elseif horarios > 19 then
						hora_ag = hora_ag + 1
					else
						hora_ag = horarios + 5
					end if
						
						esquerda = 60
						topo = topo + 40
						'coluna = 0%>
						<tr>
							<td bgcolor="silver"><%=zero(hora_ag)%>:00</td>
							<%'*** Separa e mostra cada rack ***
							unidades_array = split(conjunto_total,";")
							For Each unidades_item_junto In unidades_array
								unidades_item_junto_len = len(unidades_item_junto)
								
								separador1 = instr(unidades_item_junto,",")
								separador2 = instr(mid(unidades_item_junto,(separador1)+1,unidades_item_junto_len),",")
								
								strunidade = mid(unidades_item_junto,1,separador1-1)
								strracks = mid(unidades_item_junto,separador1+1,separador2-1)
								strqtd_racks = mid(unidades_item_junto,(separador1+separador2)+1,separador2-1)
									
									strrack_separa1 = instr(strracks,":")
									
									for i_racks = 1 to strqtd_racks '*** Nº colunas (qtd_total racks) ***
									
										if strqtd_racks > 1 then
											if i_racks <= 1 then
												strrack = mid(strracks,1,instr(strracks,":")-1)
												separador = instr(strracks,":")
											else
												strrack = mid(strracks,separador+1,instr(strracks,":")-1)
												separador = instr(strracks,":") + separador
											end if
										end if
										coluna = coluna + 1
										
										xsql = "SELECT * FROM TBL_agenda WHERE (DAY(dt_data) = "&dia_sel&") AND (MONTH(dt_data) = "&mes_sel&") AND (YEAR(dt_data) = "&ano_sel&") AND (cd_unidade = "&strunidade&") AND ({ fn HOUR(dt_data) } = "&hora_ag&") AND ({ fn MINUTE(dt_data) } BETWEEN 0 AND 59) AND (cd_rack = "&strrack&")"
										Set rs_agenda = dbconn.execute(xsql)
												
												if not rs_agenda.EOF Then
													cd_agenda = rs_agenda("cd_codigo")
													nm_medico = rs_agenda("nm_medico")
													nm_cirurgia = rs_agenda("nm_cirurgia")
													dt_tempo_estimado = rs_agenda("dt_tempo_estimado")
														tempo_estimado = dt_tempo_estimado
														cor_celula = "yellow"
												end if
												
														'*** Célula atual ***
														celula = strunidade&"-"&strrack&"+"&hora_ag&"~"&dt_tempo_estimado
														
														'*** Unidade e Rack ***
														UR = strunidade&"-"&strrack
														
														'*** Verifica se a unidade e rack estão na lista ***
														VUR = instr(1,STR_URHX,UR,1)
														
														'*** retira e trabalha a parte da strig necessária
														hora_celula = mid(celula,instr(1,celula,"+",1)+1,1)
														''hora_string = mid(STR_URHX,instr(STR_URHX,UR),instr(STR_URHX,";"))
														
														hora_string = instr(STR_URHX,UR)
															if hora_string > 0 then
														'		parte_string = mid(STR_URHX,hora_string,len(STR_URHX))
														'			hora_string_1 = instr(parte_string,"+")
														'			hora_string_2 = instr(parte_string,"~")
														'			hora_string = mid(parte_string,hora_string_1+1,hora_string_2-hora_string_1-1)
														'			dif_horas = int(hora_celula - hora_string)
																	
														'			'*** Verifica o tempo reservado a cirurgia ***
														'			X_tempo = tempo_estimado - dif_horas
														'			
														'			str_salto = str_salto  &"<br>"&UR&":"&tempo_estimado
																	
															end if
														'if 	tempo_estimado <> "" then
														'	str_salto = str_salto  &"<br>"&UR&":"&tempo_estimado
														'end if
														
														'*** Elimina a string quando o numero de hora for = 1
														'if qtd_cel_salto = 1 then
														'	str_salto = replace(str_salto,str_split_salto&";","")
														'	'str_salto = replace(str_salto,str_split_salto&";","")
														'end if
														
															'if hora_string
															
															'*** Verifica se a celula deve ser pulada ***
															'if VUR <> 0 AND str_salto <> "" then 
															'if str_salto <> ""  then 
																VSalto = Instr(str_salto,UR)
																if VSalto <> 0 then
																	localiza_limite = Instr(VSalto,str_salto,";",1)
																	str_split_salto = mid(str_salto,VSalto,localiza_limite-VSalto)
																		'*** procura o numero de saltos restantes ***
																		pos_restantes = instr(str_split_salto,":")
																		qtd_cel_salto = mid(str_split_salto,pos_restantes+1,pos_restantes-len(pos_restantes))
																			'*** Diminui o número de saltos restantes ***
																			if int(qtd_cel_salto) > 0 then
																				str_salto = replace(str_salto,str_split_salto&";",UR&":"&qtd_cel_salto-1&";")
																			end if
																end if
															'end if%>
															
											<%'if dif_horas > 0 and X_tempo > 0 Then
											'if dif_horas > 0 and X_tempo > 0 Then
												'i_racks = i_racks + dif_horas
												'pula = UR&"-1"%>
												<!--td>1. deve pular</td-->
											<%'end if%>
											
												<%'*** Elimina a string quando o numero de horas for menor que 0 ***
												if int(qtd_cel_salto) <= 0 then
													str_salto = replace(str_salto,UR&":"&qtd_cel_salto&";","")
												else
													Ver_salto = Instr(str_salto,ur)
												end if%>
												
											<%if Ver_salto = 0 then%>
												<td onDblclick="agendamento_abre('block');" align="left" valign="top" rowspan="<%=dt_tempo_estimado%>" bgcolor="<%=cor_celula%>">
													<%if cd_agenda <> "" then%>
														<!--U.<%=strunidade%> C.<%=coluna%><br>R.<%=strracks%><br>Q.<%=strqtd_racks%><br>N.<%=strrack%> S.<%=separador%>-->
														<%=nm_medico%><br><%=nm_cirurgia%><br><%=dt_tempo_estimado%><br><%=STR_URHX%>
														<%str_salto = UR&":"&tempo_estimado-1&";"&str_salto%>
													<%else%>
														.UR  = <%=UR%><br>.CEL = <%=celula%><br>.SSalto = <%=str_salto%><br>.VSalto = <%=VSalto%><%if Vsalto <> 0 then response.write("<b>OK</b>") end if%><br>.Ver_Slt = <%'=Ver_salto%><%if Ver_salto <> 0 then%><b>OK</b><%end if%><br>.Spl = <%=str_split_salto%><br>.QTD_slt = <%=qtd_cel_salto%><br>.teste=<%=str_salto%>-->
														<a href="javascript: void(0);"  onClick="agendamento_abre('block')"><img src="../imagens/px.gif" alt="" width="85" height="1" border="0"></a>
													<%end if%>
												</td>
											<%end if%>												
										<%if dt_tempo_estimado <> "" then
											STR_URHX = STR_URHX&";"&strunidade&"-"&strrack&"+"&hora_ag&"~"&dt_tempo_estimado
										end if
										
										
										
										'if dt_tempo_estimado < "0" and then
										'	i_racks = i_racks + 1
										'end if
										col_dif = 1
										%>
										
										
										<%cd_agenda = ""
										nm_medico = ""
										nm_cirurgia = ""
										dt_tempo_estimado = ""
										tempo_estimado = ""
										dif_horas = 0
										pula = UR&"-0"
										VSalto = 0
										localiza_limite = 0
										str_split_salto = 0
										
										cor_celula = "White"
									
									next
									'strrack = ""
									
							next
							coluna = "0"%>
							<td><%=zero(hora_ag)%>:00</td>
						</tr>
				<%next%>
				
				<tr>
					<td>&nbsp;</td>
					<td colspan="<%=qtd_racks_ativos%>">&nbsp;<%=qtd_racks_ativos%></td>
					<td>&nbsp;</td>
				</tr>
				</form>
			</table>
			<br>
			<br>
			

<div align="center">
	<form action="agenda/acoes/agenda_acoes.asp" method="post">
		<input type="hidden" name="cd_situacao" value="<%=str_situacao%>">
	<table id="nome" style="position:absolute; left:15px; top:20px; background-color:silver; z-index=-1;">
		<tr><td colspan="2" align="center" style="border:1px solid black;"><a href="javascript: void(0);" onMouseOver="mostra_campos();">Mostra</a></td></tr-->
		<tr>
			<td>Hospital:</td>
			<td>
				<%xsql ="SELECT * FROM TBL_unidades where cd_status=1"
					Set rs = dbconn.execute(xsql)%>
					<select name="cd_unidade">
						<option value=""></option>
					<%do while Not rs.EOF
					cd_unidade = rs("cd_codigo")
					nm_unidade = rs("nm_unidade")%>
					
						<option value="<%=cd_unidade%>"><%=nm_unidade%></option>
					
					<%=unidade%>
					<%rs.movenext
					loop%></select></td>
		</tr>
		<tr>
			<td>Nº Reserva:</td><td><input type="text" name="cd_reserva" size="20" value="<%=ltrim(str_reserva)%>" style="font-size:9;"> &nbsp; <i style="color:red;"><%=str_situacao%></i></td>
		</tr>
		<tr>	
			<td>Paciente:</td><td><input type="text" name="nm_paciente" size="40" value="<%=ltrim(str_paciente)%>" style="font-size:9;"></td>
		</tr>
		<tr>
			<td>Data Cir.:</td><td><input type="text" name="dt_cirurgia" size="20" value="<%=ltrim(str_datacirurg)%>" style="font-size:9;"></td>
		</tr>
		<tr>
			<td>Médico:</td><td><input type="text" name="nm_medico" size="40" value="<%=ltrim(str_medico)%>" style="font-size:9;"></td>
		</tr>
		<tr>
			<td>Centro:</td><td><input type="text" name="nm_centro" size="40" value="<%=ltrim(mid(str_centro,1,len(str_centro)))%>" style="font-size:9;"></td>
		</tr>
		
		<tr>
			<td>Cirurgia:</td><td><input type="text" name="nm_cirurgia" size="40" value="<%=str_amb%><%'=str_cirurgia%>" style="font-size:9;"></td>
		</tr>
		<tr>
			<%str_materiais = replace(str_materiais,vbcrlf,"")
			str_materiais = replace(str_materiais,"Material  :",";"&vbcrlf)
			str_materiais = replace(str_materiais,"Fornecedor:",",")
			str_materiais = replace(str_materiais,"Quantidade:",",")%>
			<td>Material:</td>
			<td><textarea cols="40" rows="3" name="nm_material" style="font-size:9;"><%=str_materiais%></textarea>			
		</tr>
		<tr>
			<%str_equipamentos = ltrim(replace(str_equipamentos,vbcrlf,"#"))
			str_equipamentos = ltrim(replace(str_equipamentos,"##",""))
			str_equipamentos = ltrim(replace(str_equipamentos,"#",vbcrlf))
			'str_equipamentos = replace(str_equipamentos,"Equip.:","<b>EQUIP.:</b>")
			'str_equipamentos = replace(str_equipamentos,"Equip.:",",")
			'str_equipamentos = replace(str_equipamentos,"Caixa.:","")%>
			<td>Equip.:</td>
			<td><textarea cols="40" rows="3" name="nm_equipamentos" style="font-size:9;"><%=str_equipamentos%></textarea></td>	
		</tr>
		<tr>
			<td>Obs.:</td>
			<td><textarea cols="40" rows="3" name="nm_obs" style="font-size:9;"><%=str_obs%></textarea></td>
		</tr>
		<tr><td align="right" ><a href="javascript: void(0);" onMouseOver="esconde_campos();"><b>X</b></a></td><td align="center"><input type="submit" name="ok" value="OK">  <%=nr_reserva%></td></tr>
	</table>
	</form>
	
	<!--div id="flutua" idis="campo_agendamento" style="position:absolute; top:200px; left: 400px; display:none;"-->
	<div id="flutua" style="border: 1px solid red; background: #e1e2f3; padding: 20px; display: none; position: absolute">
		<table style="border:1px solid red; background-color:silver;">
			<tr>				
				<td>teste de posicionamento de janela: <a href="javascript: void(0);"  onclick="agendamento_abre('none')">Fecha</a><form action="#" name="flutuador" id="flutuador"><input type="text" name="flutua_pos"></form></td>
			</tr>
			<tr><td><input type="text" name="teste"></td></tr>
		</table>
	</div>
	<div id="mostra_agendamento" style="position:absolute; top:200px; left: 400px; display:none;">
		<span id='divAgendamento'></span>		
	</div>
	<br>
	<br>
	
</div>
<div id="flutua2" style="border: 1px solid red; background: #e1e2f3; padding: 20px; display: none; position: absolute">kaduzick</div>
<label onmouseover="mostra(event)" >teste</label>
</body>
</html>

<!--
DATA RESERVA  
DATA INTERNAC.
DATA CIRURGIA 
CIRURGIA      
MEDICO        
MEDICO CREDENCIADO NO CONVENIO?
PATOLOGISTA   
PROTESES      
CENTRO   
AMB
Material  
Fornecedor
Quantidade
Material  
Fornecedor
Quantidade

-->







</body>
</html>
