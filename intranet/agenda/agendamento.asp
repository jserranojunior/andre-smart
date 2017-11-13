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
	
	function mostra_agendamento(situacao,aidi,campo) {		
		document.getElementById('mostra_agendamento').style.display=situacao;
		//Pega a posição do elemento no eixo X
		var obj_x = document.getElementById(aidi);		
		var curleft = 0;
		if(obj_x.offsetParent)
			while(1){
				curleft += obj_x.offsetLeft;
				if(!obj_x.offsetParent)
					break;
				obj_x = obj_x.offsetParent;
			}
		else if(obj_x.x)
			curleft += obj_x.x;
		document.getElementById('mostra_agendamento').style.left=curleft-210;
		
		//Pega a posição do elemento no eixo Y
		var obj_y = document.getElementById(aidi);			
		var curtop = 0;
		if(obj_y.offsetParent)
			while(1){
				curtop += obj_y.offsetTop;
				if(!obj_y.offsetParent)
					break;
				obj_y = obj_y.offsetParent;
			}
		else if(obj_y.y)
			curtop += obj_y.y;
		document.getElementById('mostra_agendamento').style.top=curtop-21;
		
		

		setTimeout('document.getElementById("mostra_agendamento").style.display="none"',18000);
		
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
<!--#include file="../js/mouseposition.js"-->
<body onLoad="inicio()">

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
				response.write("Não foi possível determinar qual a unidade.<br>Por favor, preencha o formulário.")
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


<!--div id="mostra_agendamento" style="position:relative; top:1px; left: 1px; display:none;">
		<span id='divAgendamento'></span>		
	</div-->

<%'=informacao_bruta%>

		
			<!-- onClick="controlaCamada('nome')"-->
			
			<table border="2" style="border-collapse=collapse;" onMouseOver="esconde_campos();">
				<form action="agendamento.asp" name="Form" id="Form">
				<input type="hidden" name="ajax">
				<input type="text" name="posx">
				<tr>
					<td colspan="100%" align="center" style="font-size:16px;">AGENDA</td>
				</tr>
				<tr>
					<td colspan="100%" align="center">Hoje é: <%=DiaSemanaExtenso(weekday(dia_sel&"/"&mes_sel&"/"&ano_sel))%>, <%=zero(dia_sel)%> de <%=mes_selecionado(mes_sel)%> de <%=ano_sel%></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<%x = 1
					
					'**********************************
					'*** A - Conta as unidades ativas ***
					'************************************
					xsql = "SELECT distinct(cd_ordem_unidade) AS cd_unidades_ativas,cd_unidade,nm_unidade,nm_sigla,nm_cor FROM VIEW_agenda_unidade_rack_ordem WHERE (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011')  ORDER BY cd_ordem_unidade"
					Set rs_und = dbconn.execute(xsql)
					while not rs_und.EOF
						cd_unidade = rs_und("cd_unidade")
						nm_unidade = rs_und("nm_unidade")
						nm_sigla = rs_und("nm_sigla")
						nm_cor = rs_und("nm_cor")
						qtd_unidades_ativas = int(qtd_unidades_ativas + 1)
						
							'**********************************************************************************
							'*** B - Separa as unidades, os racks e conta quantos racks cada unidade possui ***
							'**********************************************************************************
							xsql = "SELECT * FROM VIEW_agenda_unidade_rack_ordem WHERE cd_unidade = "&cd_unidade&" AND (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011') ORDER BY cd_ordem_unidade, cd_ordem_rack"
							'xsql = "SELECT * FROM VIEW_unidades_racks WHERE cd_unidade = "&cd_unidade&" AND (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato IS NULL) OR (cd_hospital = 1) AND (dt_inicio_contrato >= '01/01/2000') AND (dt_fim_contrato <= '12/31/2011') ORDER BY cd_ordem_unidade, cd_ordem_rack"
							
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
									<td align="center" colspan="<%=qtd_racks_unidade%>" style="border:2px solid #5c5c5c; background-color:#<%=nm_cor%>;"><%=nm_unidade%><%'=qtd_racks_unidade%></td>
								<%else%>
									<td align="center" colspan="<%=qtd_racks_unidade%>" style="border:2px solid #5c5c5c; background-color:#<%=nm_cor%>;"><%=nm_sigla%><%'=qtd_racks_unidade%></td>
								<%end if%>
						<%'*** Grava a string com as unidades, racks e quantidades de racks ***
						str_URQ = str_URQ &";("&cd_unidade&")"&racks_unidades&":#"&qtd_racks_unidade
						
						racks_unidades = ""
						qtd_racks_unidade = "0"
					rs_und.movenext
					wend
					
					str_URQ = mid(str_URQ,2,len(str_URQ))%>
					<td>&nbsp;</td>
				</tr>
				
				<tr>
					<td>&nbsp;</td>
					<td colspan="15">&nbsp;<%'=str_URQ%></td>
				</tr>
				<tr>
					<td style="border-right:solid 2px #5c5c5c;">&nbsp;</td><!-- *** RACKS *** -->
					<%'***************************************************
					'*** D - Separa e mostra os racks de cada unidade ***
					'*****************************************************
					UR_array = split(str_URQ,";")
					For Each str_URQ_junto In UR_array
						str_URQ_junto_len = len(str_URQ_junto)
						separador_Q = instr(str_URQ_junto,"#")
						'cd_unidade = 
					
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
									cd_rack = mid(str_rack_parc,1,instr(str_rack_parc,":")-1)
								'*** Mostra a Unidade ***
									cd_unidade = mid(str_URQ_junto,2,instr(str_URQ_junto,")")-2)
								'*** Separa as colunas das unidades (racks) ***
									if int(qtd_Racks) = int(Ri) then
										borda_direita = "border-right:solid 2px #5c5c5c;"
									end if
										'*** Mostra o nome do Rack ***
										xsql = "SELECT * FROM TBL_unidades_racks WHERE cd_unidade = "&cd_unidade&" AND cd_rack = "&cd_rack&""
										Set rs_rack = dbconn.execute(xsql)
										while not rs_rack.EOF
											nm_rack = rs_rack("nm_rack")
											cd_sala = rs_rack("cd_sala")
												If cd_sala <> "" Then
													cd_sala = cd_sala&"<br>"
												end if
										rs_rack.movenext
										wend%>
									
										<td valign="bottom" align="center" style="border-top:solid 2px #5c5c5c; <%=borda_direita%>"><b><%=cd_sala%><%=nm_rack%></b></td>
								
							<%str_rack_parc = mid(str_rack_parc,pos_separador+1,len(str_rack_parc))
							pos_separador_ult = pos_separador
							
							next
							cd_rack = 0
							pos_separador_ult = ""
							borda_direita = ""
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
						%>
						<tr>
							<td style="background-color:silver; color:white; border-right:solid 2px #5c5c5c;"><%=zero(hora_ag)%>:00</td>
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
												
												'*** Separa as colunas das unidades (racks) ***
												if int(qtd_Racks) = int(Ri) then
													borda_direita = "border-right:solid 2px #5c5c5c;"
												end if
												if int(hora_ag) = 5 then
													borda_inferior = "border-bottom:solid 2px #5c5c5c;"
												end if
											'*** Mostra o Rack ***
												cd_rack = mid(str_rack_parc,1,instr(str_rack_parc,":")-1)%>
													<%'**************************
													'*** Mostra o agendamento ***
													'****************************
													cor_celula = "#ffffff"
													str_verifica = 0
													
													xsql = "SELECT * FROM TBL_agenda WHERE (DAY(dt_data) = "&dia_sel&") AND (MONTH(dt_data) = "&mes_sel&") AND (YEAR(dt_data) = "&ano_sel&") AND (cd_unidade = "&cd_unidade&") AND ({ fn HOUR(dt_data) } = "&hora_ag&") AND ({ fn MINUTE(dt_data) } BETWEEN 0 AND 59) AND (cd_rack = "&cd_rack&")"
													Set rs_agenda = dbconn.execute(xsql)
														if not rs_agenda.EOF Then
															cd_agenda = rs_agenda("cd_codigo")
															nm_medico = rs_agenda("nm_medico")
															nm_cirurgia = rs_agenda("nm_cirurgia")
															nm_obs = rs_agenda("nm_obs")
															nm_aviso = rs_agenda("nm_aviso")
																if nm_aviso <> "" Then
																	nm_aviso = "<br><span style='color:#ff0033; font-size:12px; font-weight:bold;'>"&nm_aviso&"</span>"
																end if
															
															dt_hora = zero(hour(rs_agenda("dt_data")))&":"&zero(minute(rs_agenda("dt_data")))
															dt_hora_ag = zero(hora_ag)&":00"
																if dt_hora <> dt_hora_ag then
																	mostra_hora = "<span style='color:#990099; font-size:12px; font-weight:bold;'>"&dt_hora&"</span><br>"
																end if
															
															dt_tempo_estimado = rs_agenda("dt_tempo_estimado")
																tempo_estimado = dt_tempo_estimado
																
															qtd_cir = qtd_cir + 1
															
															qtd_cir_agendadas = qtd_cir_agendadas + 1
															
															bordas = "border:solid 2px #ff6600;"
															
																'tempoE_teste = tempo_estimado
																cor_celula = "#e5e5e5"
																str_verifica = 1
																
																'***  ***
																'*** verifica os tamanhos das palavras para determinar a largura da célula ***
																len_med = len(nm_medico)
																len_cir = len(nm_cirurgia)
																len_obs = len(nm_obs)
																
																if len_med > len_cir OR len_med > len_obs then
																	larg_cel = len_med * 5
																elseif len_cir > len_med OR len_cir > len_obs then
																	larg_cel = len_cir * 5
																elseif len_obs > len_med OR len_obs > len_cir then
																	larg_cel = len_obs * 5
																else
																	larg_cel = 45
																end if
															
															'*** grava sting para esconder celula atingida pelo rowspan ***
															if str_esconde_celula = "" Then
																str_esconde_celula = "["&cd_unidade&"]"&cd_rack&"!"&tempo_estimado&"*"&hora_ag&"/"
															else
																str_esconde_celula = str_esconde_celula&";["&cd_unidade&"]"&cd_rack&"!"&tempo_estimado&"*"&hora_ag&"/"
															end if
														Else
															larg_cel = 0
														end if
														
														'*** subtrai 1 hora na string: "str_esconde_celula", referente a celula atingida pelo rowspan ***
														if instr(str_esconde_celula,"["&cd_unidade&"]"&cd_rack&"!") <> 0 then
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
																		if instr(str_esconde_celula,str_celula_split_UR) <> 0 then
																			str_verifica = int(str_verifica) + 1
																				'*** Diminui a hora estimada e aumenta a hora agendada ***
																				str_esconde_celula = replace(str_esconde_celula,"["&str_celula_U&"]"&str_celula_R&"!"&str_celula_T&"*"&str_celula_H&"/","["&str_celula_U&"]"&str_celula_R&"!"&str_celula_T-1&"*"&str_celula_H+1&"/")
																				if  str_verifica = 1 Then
																					salta = 1
																				end if
																					'*** Apaga a string caso a hora estimada seja igual a 1 ***
																					if str_celula_T = 1 then
																						str_esconde_celula = replace(str_esconde_celula,"["&str_celula_U&"]"&str_celula_R&"!"&str_celula_T-1&"*"&str_celula_H+1&"/","")
																					end if
																		else
																			salta = 0
																		end if
																	
														end if
														
														if salta = 0 then%>
														<td valign="top" align="center" rowspan="<%=dt_tempo_estimado%>" style="background-color:<%=cor_celula%>; font-size:9px; <%=borda_direita%> <%=borda_inferior%> <%=bordas%>" <%if cd_agenda <> "" Then%>id="elemento<%=x%>" onmouseover="mostra_agendamento('block','elemento<%=x%>','<%=cd_agenda%>');" onclick="mostra_agendamento('block','elemento<%=x%>','<%=cd_agenda%>');" <%end if%>>
															
															<%if cd_agenda <> "" Then%>
																<!--a href="javascript: void(0);" onMouseover="mostra_agendamento('block',<%=topo%>,<%=esquerda%>,'elemento<%=x%>','<%=cd_agenda%>');"-->
																	<%=mostra_hora%>
																	<%=nm_medico%><br>
																	<%=nm_cirurgia%><br>
																	<%=nm_obs%>
																	<%=nm_aviso%>
																	<img src="../imagens/px.gif" alt="" width="<%=larg_cel%>" height="1" border="0">
																	<!--<%'if cd_agenda <> "" Then%>
																	<input type="checkbox" name="confirmacao" value="">																	
																	<%'end if%>-->
																	<input type="hidden" name="<%=cd_agenda%>" value="<%=cd_unidade%>,<%=cd_rack%>,<%=zero(hora_ag)%>:00" size="5">
																</a>
																<!--input type="text" name="x" value="elemento<%'=x%>"  size="8" -->
															<%x = x + 1
															end if%>
														</td>
														<%end if
														
													
													
													cd_agenda = ""
													nm_medico = ""
													nm_cirurgia = ""
													nm_obs = ""
													dt_hora = ""
													nm_aviso = ""
													
													dt_tempo_estimado = ""
													str_esconde_celula_teste = 0
													inicio_str = ""
													str_verifica = ""
													mostra_hora = ""
													larg_cel = ""
													borda_direita = ""
													bordas = ""
													
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
					<%'***************************************************
					'*** D - Separa e mostra os racks de cada unidade ***
					'*****************************************************
					UR_array = split(str_URQ,";")
					For Each str_URQ_junto In UR_array
						str_URQ_junto_len = len(str_URQ_junto)
						separador_Q = instr(str_URQ_junto,"#")
						'cd_unidade = 
					
						qtd_Racks = mid(str_URQ_junto,separador_Q+1,str_URQ_junto_len-separador_Q)
						
						'*** 
						str_Rack = mid(str_URQ_junto,instr(str_URQ_junto,":")+1,instr(str_URQ_junto,"#"))
							str_Rack = mid(str_Rack,1,instr(str_Rack,"#")-1)
							str_rack_parc = str_Rack
							
							'*** Separa os Racks ***
								pos_separador = instr(str_rack_parc,":")
							'for Ri=1 to qtd_Racks
								str_rack_parc = mid(str_rack_parc,1,len(str_rack_parc))
								'*** Mostra o Rack ***
									cd_rack = mid(str_rack_parc,1,instr(str_rack_parc,":")-1)
								'*** Mostra a Unidade ***
									cd_unidade = mid(str_URQ_junto,2,instr(str_URQ_junto,")")-2)
								'*** Separa as colunas das unidades (racks) ***
									if int(qtd_Racks) = int(Ri) then
										borda_direita = "border-right:solid 2px #5c5c5c;"
									end if
										'*** Mostra o nome do Rack ***
										'xsql = "SELECT * FROM TBL_unidades_racks WHERE cd_unidade = "&cd_unidade&" AND cd_rack = "&cd_rack&""
										xsql = "up_agenda_qtd_cir_unidade @cd_unidade="&cd_unidade&", @mes_sel="&mes_sel&", @ano_sel="&ano_sel&", @dia_sel="&dia_sel&""
										Set rs_q_cir = dbconn.execute(xsql)
										if not rs_q_cir.EOF then
											qtd_cir = rs_q_cir("qtd_cir")
										rs_q_cir.movenext
										end if
										
										%>
									
										<td valign="bottom" align="right" colspan="<%=qtd_Racks%>" style="border:2px solid #5c5c5c; <%=borda_direita%>"><b><%=qtd_cir%></b>&nbsp;</td>
								
							<%str_rack_parc = mid(str_rack_parc,pos_separador+1,len(str_rack_parc))
							pos_separador_ult = pos_separador
							
							'next
							cd_rack = 0
							pos_separador_ult = ""
							borda_direita = ""
							%>
					<%next%>
					<td style="font-size:18px; font-weigth:bold;">&nbsp;<%=qtd_cir_agendadas%></td>
				</tr>
				</form>
			</table>
			
			
	<div id="mostra_agendamento" style="position:absolute; display:none;">
		<span id='divAgendamento'></span>		
	</div>			
			

<div align="center">
	<form action="agenda/acoes/agenda_acoes.asp" method="post" name="ficha">
		<input type="hidden" name="cd_situacao" value="<%=str_situacao%>">
			
				
	<table id="nome" style=" border-collapse:collapse; border=0px solid gray; position:absolute; left:15px; top:20px; background-color:#dddddd; z-index=-1;" onMouseOver="mostra_campos();">
	<!--table id="nome" style=" border-collapse:collapse; border=0px solid gray; position:absolute; background-color:#dddddd; z-index=-1;" onMouseOver="mostra_campos();"-->
		<tr>
			<td colspan="2" align="center" style="border:1px solid black;"><b>AGENDAMENTO</b></td>
			<td rowspan="13"><img src="../imagens/sombra_v_inicio.gif" alt="" width="20" height="22" border="0"><br><img src="../imagens/sombra_v_meio.gif" alt="" width="20" height="400" border="0"></td>
		</tr>
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
		<tr>
			<td style="color:red;">Alerta:</td>
			<td><textarea cols="40" rows="3" name="nm_alerta" style="font-size:9; color:red;"><%=str_alerta%></textarea></td>
		</tr>
		<tr><td align="center" colspan="2"><input type="submit" name="ok" value="OK">  <%=nr_reserva%></td></tr>
		<tr>
			<td colspan="3"><img src="../imagens/sombra_h_inicio.gif" alt="" width="20" height="20" border="0"><img src="../imagens/sombra_h_meio.gif" alt="" width="300" height="20" border="0"><img src="../imagens/sombra_hv_canto.gif" alt="" width="20" height="20" border="0"></td>
		</tr>
	</table>			
	</form>
</div>

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

<script language="javascript">
	function inicio() {
		if (document.ficha.cd_reserva.value != ""){
			document.getElementById('nome').style.zIndex="1";
		}
	}
	
	
	
	window.onload=inicio;
</script>



</body>
</html>
