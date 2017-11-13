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
				response.write("Não foi possível determinar qual a unidade.<br>Por favor, preencha o formulário abaixo.")
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
					<td colspan="100%" align="center" style="font-size:16px;">AGENDA</td>
				</tr>
				<tr>
					<td colspan="100%" align="center">Hoje é: <%=zero(dia_sel)%> de <%=mes_selecionado(mes_sel)%> de <%=ano_sel%></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<%'**********************************
					'*** A - Conta as unidades ativas ***
					'************************************
					xsql = "SELECT distinct(cd_ordem_unidade) AS cd_unidades_ativas,cd_unidade,nm_unidade,nm_sigla FROM VIEW_agenda_unidade_rack_ordem WHERE (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011')  ORDER BY cd_ordem_unidade"
					Set rs_und = dbconn.execute(xsql)
					while not rs_und.EOF
						cd_unidade = rs_und("cd_unidade")
						nm_unidade = rs_und("nm_unidade")
						nm_sigla = rs_und("nm_sigla")
						qtd_unidades_ativas = int(qtd_unidades_ativas + 1)
						
							'**********************************************************************************
							'*** B - Separa as unidades, os racks e conta quantos racks cada unidade possui ***
							'**********************************************************************************
							xsql = "SELECT * FROM VIEW_agenda_unidade_rack_ordem WHERE cd_unidade = "&cd_unidade&" AND (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011') ORDER BY cd_ordem_unidade, cd_ordem_rack"
							Set rs_und_rack = dbconn.execute(xsql)
							while not rs_und_rack.EOF
								cd_rack = rs_und_rack("cd_rack")
								qtd_racks_unidade = qtd_racks_unidade + 1
								
									'*** Cria string com os racks da unidade referente ***
									racks_unidades = racks_unidades + ":"&cd_rack
								
							rs_und_rack.movenext
							wend
								
								'****************************
								'*** C - Mostra a unidade ***
								'****************************
								if qtd_racks_unidade >= 3 then%>
									<td colspan="<%=qtd_racks_unidade%>"><%=nm_unidade%>(<%=qtd_racks_unidade%>)</td>
								<%else%>
									<td colspan="<%=qtd_racks_unidade%>"><%=nm_sigla%>(<%=qtd_racks_unidade%>)</td>
								<%end if%>
						<%'*** Grava a string com as unidades, racks e quantidades de racks ***
						str_URQ = str_URQ &";("&cd_unidade&")"&racks_unidades&":#"&qtd_racks_unidade
						
						racks_unidades = ""
						qtd_racks_unidade = "0"
					rs_und.movenext
					wend
					
					str_URQ = mid(str_URQ,2,len(str_URQ))%>
				</tr>
				
				<tr>
					<td>&nbsp;</td>
					<td colspan="15"><%=str_URQ%></td>
				</tr>
				<tr>
					<td>&nbsp;</td><!-- *** RACKS *** -->
					<%'***************************************************
					'*** D - Separa e mostra os racks de cada unidade ***
					'*****************************************************
					UR_array = split(str_URQ,";")
					For Each str_URQ_junto In UR_array
						str_URQ_junto_len = len(str_URQ_junto)
						separador_Q = instr(str_URQ_junto,"#")
					
						qtd_Racks = mid(str_URQ_junto,separador_Q+1,str_URQ_junto_len-separador_Q)
						
						'*** 
						str_Rack = mid(str_URQ_junto,instr(str_URQ_junto,":")+1,instr(str_URQ_junto,"#"))
							str_Rack = mid(str_Rack,1,instr(str_Rack,"#")-1)
							str_rack_parc = str_Rack
							
							'*** Separa os Racks ***
								pos_separador = instr(str_rack_parc,":")
							for Ri=1 to qtd_Racks
								str_rack_parc = mid(str_rack_parc,1,len(str_rack_parc))
								'*** Mostra o Rack ***
									cd_rack = mid(str_rack_parc,1,instr(str_rack_parc,":")-1)%>
									
										<td valign="top" align="center"><b><%=cd_rack%></b></td>
								
							<%str_rack_parc = mid(str_rack_parc,pos_separador+1,len(str_rack_parc))
							pos_separador_ult = pos_separador
							
							next
							cd_rack = 0
							pos_separador_ult = ""
							%>
					<%next%>
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
							<td style="background-color:silver; color:white;"><%=zero(hora_ag)%>:00</td>
								<%'***************************************
								'*** E - Mostra as cirurgias agendadas ***
								'*****************************************
								
								'*** E1. Separa os racks de cada unidade ***
								UR_array = split(str_URQ,";")
								For Each str_URQ_junto In UR_array
									str_URQ_junto_len = len(str_URQ_junto)
									separador_Q = instr(str_URQ_junto,"#")
									
									cd_unidade = mid(str_URQ_junto,2,instr(str_URQ_junto,")")-2)
									qtd_Racks = mid(str_URQ_junto,separador_Q+1,str_URQ_junto_len-separador_Q)
									
									'*** 
									str_Rack = mid(str_URQ_junto,instr(str_URQ_junto,":")+1,instr(str_URQ_junto,"#"))
										str_Rack = mid(str_Rack,1,instr(str_Rack,"#")-1)
										str_rack_parc = str_Rack
										
										'*** Separa os Racks ***
											pos_separador = instr(str_rack_parc,":")
										for Ri=1 to qtd_Racks
											str_rack_parc = mid(str_rack_parc,1,len(str_rack_parc))
											'*** Mostra o Rack ***
												cd_rack = mid(str_rack_parc,1,instr(str_rack_parc,":")-1)%>
													<%'**************************
													'*** Mostra o agendamento ***
													'****************************
													cor_celula = "white"
													str_verifica = 0
													
													xsql = "SELECT * FROM TBL_agenda WHERE (DAY(dt_data) = "&dia_sel&") AND (MONTH(dt_data) = "&mes_sel&") AND (YEAR(dt_data) = "&ano_sel&") AND (cd_unidade = "&cd_unidade&") AND ({ fn HOUR(dt_data) } = "&hora_ag&") AND ({ fn MINUTE(dt_data) } BETWEEN 0 AND 59) AND (cd_rack = "&cd_rack&")"
													Set rs_agenda = dbconn.execute(xsql)
														if not rs_agenda.EOF Then
															cd_agenda = rs_agenda("cd_codigo")
															nm_medico = rs_agenda("nm_medico")
															nm_cirurgia = rs_agenda("nm_cirurgia")
															dt_tempo_estimado = rs_agenda("dt_tempo_estimado")
																tempo_estimado = dt_tempo_estimado
																'tempoE_teste = tempo_estimado
																cor_celula = "yellow"
																str_verifica = 1
															
															'*** grava sting para esconder celula atingida pelo rowspan ***
															if str_esconde_celula = "" Then
																str_esconde_celula = "["&cd_unidade&"]"&cd_rack&"!"&tempo_estimado&"*"&hora_ag&"/"
															else
																str_esconde_celula = str_esconde_celula&";["&cd_unidade&"]"&cd_rack&"!"&tempo_estimado&"*"&hora_ag&"/"
															end if
														end if
														
														'*** subtrai 1 hora na string: "str_esconde_celula", referente a celula atingida pelo rowspan ***
														if instr(str_esconde_celula,"["&cd_unidade&"]"&cd_rack&"!") <> 0 then
															'str_esconde_celula_teste = 1
																'*** Localiza a posição da celula na string ***
																inicio_str = instr(str_esconde_celula,"["&cd_unidade&"]"&cd_rack&"!")
																str_celula_split = mid(str_esconde_celula,inicio_str,len(str_esconde_celula))
																	'*** identifica o caractere "/" (final) ***
																	str_celula_split_fim = instr(str_celula_split,"/")
																	str_celula_split = mid(str_celula_split,1,str_celula_split_fim)
																		
																		
																		str_celula_split_UR = mid(str_celula_split,1,instr(str_celula_split,"!"))
																		'*** Identifica os demais caracteres referentes as separações ***
																		str_celula_split_col2 = instr(str_celula_split,"]")
																		str_celula_split_escl = instr(str_celula_split,"!")
																		str_celula_split_ast = instr(str_celula_split,"*")
																		str_celula_split_bar = instr(str_celula_split,"/")
																		
																		str_celula_U = mid(str_celula_split,2,instr(str_celula_split_UR,"]")-2)
																		str_celula_R = mid(str_celula_split,str_celula_split_col2+1,(str_celula_split_escl - str_celula_split_col2)-1)
																		str_celula_T = mid(str_celula_split,str_celula_split_escl+1,(str_celula_split_ast - str_celula_split_escl)-1)
																		str_celula_H = mid(str_celula_split,str_celula_split_ast+1,(str_celula_split_bar-str_celula_split_ast)-1)
																		
																		'*** Verifica se a coluna encontra-se na string de agendamento ***
																		'if instr(str_esconde_celula,str_celula_split) <> 0 then
																		if instr(str_esconde_celula,str_celula_split_UR) <> 0 then
																			'str_esconde_celula = 
																			str_verifica = int(str_verifica) + 1
																				'*** Diminui a hora estimada e aumenta a hora agendada ***
																				str_esconde_celula = replace(str_esconde_celula,"["&str_celula_U&"]"&str_celula_R&"!"&str_celula_T&"*"&str_celula_H&"/","["&str_celula_U&"]"&str_celula_R&"!"&str_celula_T-1&"*"&str_celula_H+1&"/")
																				'cor_celula = "yellow"
																				if  str_verifica = 1 Then
																					salta = 1
																				end if
																					'*** Apaga a string caso a hora estimada seja igual a 1 ***
																					if str_celula_T = 1 then
																						str_esconde_celula = replace(str_esconde_celula,"["&str_celula_U&"]"&str_celula_R&"!"&str_celula_T-1&"*"&str_celula_H+1&"/","")
																						'str_verifica = ""
																						'salta = 1
																					end if
																		else
																			salta = 0
																		end if
																	
																	
																'	str_tempo_est = mid(str_esconde_celula,inicio_str
														end if
														
														if salta = 0 then%>
														<td valign="top" align="center" rowspan="<%=dt_tempo_estimado%>" style="background-color:<%=cor_celula%>; font-size:8px;">
															<%=nm_medico%><br>
															<%=cd_unidade%>-<%=cd_rack%>-<%=dt_tempo_estimado%><br>
															<%=str_esconde_celula%><br>
															<br>
															<%=str_celula_split%><br><br>
															<%=str_celula_split_UR%><br>
															<br>
															Col=<%=str_celula_split_col2%><br>
															Esc=<%=str_celula_split_escl%><br>
															Ast=<%=str_celula_split_ast%><br>
															Bar=<%=str_celula_split_bar%><br>
															<br>
															
															U=<%=str_celula_U%><br>
															R=<%=str_celula_R%><br>
															T=<%=str_celula_T%><br>
															H=<%=str_celula_H%><br>
															<br>
																														
															<%=str_verifica%>
															<img src="../imagens/px.gif" alt="" width="55" height="1" border="0">
														</td>
														<%end if
														'if instr(str_esconde_celula,str_celula_split_UR) <> 0 then
														'	salta = 1
														'end if
													nm_medico = ""
													dt_tempo_estimado = ""
													str_esconde_celula_teste = 0
													inicio_str = ""
													str_verifica = ""
													
													str_celula_U = ""
													str_celula_R = ""
													str_celula_T = ""
													salta = "0"
													%>
													
										<%str_rack_parc = mid(str_rack_parc,pos_separador+1,len(str_rack_parc))
										pos_separador_ult = pos_separador
										
										next
										cd_rack = 0
										pos_separador_ult = ""
										%>
								<%next%>
							<td style="background-color:silver; color:white;"><%=zero(hora_ag)%>:00</td>
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
