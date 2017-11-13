<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Agenda</title>
<script type="text/javascript"> 
	function pasteClipboard() { 
		//pasteClipboard função () ( 
			txt = document.form.informacao_bruta.createTextRange();
			//txt = document.form.informacao_bruta.createTextRange ();
			txt.execCommand('Paste');
			//txt.execCommand ('colar');
		} 
		//)
	
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
		setTimeout('document.getElementById("mostra_agendamento").style.display="none"',8000);
		
			//for (i=0; i<5000; i++){
			document.Form.ajax.value=campo;
			{	
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
<body>
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
					<td colspan="100%" align="center">Agenda</td>
				</tr>
				<tr>
					<td colspan="100%" align="center">Hoje é: <%=zero(dia_sel)%> de <%=mes_selecionado(mes_sel)%> de <%=ano_sel%></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				<%xsql ="SELECT COUNT(cd_rack) AS qtd_racks,cd_unidade,nm_unidade,cd_ordem FROM VIEW_unidades_racks WHERE cd_hospital=1 GROUP BY cd_unidade,nm_unidade,cd_ordem ORDER BY cd_ordem"
						'SELECT     * FROM         VIEW_unidades_racks WHERE     (cd_hospital = 1)
				
					Set rs_racks = dbconn.execute(xsql)
					while Not rs_racks.EOF
						qtd_racks = rs_racks("qtd_racks")
						cd_unidade = rs_racks("cd_unidade")
						nm_unidade = rs_racks("nm_unidade")%>
						<td colspan="<%=qtd_racks%>" align="center"><%'=cd_unidade%><%=nm_unidade%></td>
					<%racks_total = racks_total&";"&cd_unidade&","&qtd_racks
					rs_racks.movenext
					wend%>
					<td>&nbsp;</td>
				</tr>
				<%topo = 150
				nome_campo = 0
				for horarios = 1 to 24
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
							<td><%=zero(hora_ag)%>:00</td>
							<%span_total = mid(racks_total,2,len(racks_total))
							rack_array = split(span_total,";")
							For Each rack_junto In rack_array
								separador1 = int(instr(rack_junto,","))
								
									cod_unidade = mid(rack_junto,1,separador1-1)
									total_cel = mid(rack_junto,(separador1)+1,len(rack_junto))
										for rack = 1 to total_cel
											esquerda = esquerda + 35%>
											<!--td onclick="javascript:alert('rack:<%=rack%> (<%=zero(hora_ag)%>:00) Unid:<%=cod_unidade%>')"-->
											<td onDblclick="agendamento_abre('block');">
												<%'%xsql ="SELECT COUNT(cd_rack) AS qtd_racks,cd_unidade,nm_unidade,cd_ordem FROM VIEW_unidades_racks WHERE cd_hospital=1 GROUP BY cd_unidade,nm_unidade,cd_ordem ORDER BY cd_ordem"
														'SELECT     * FROM         VIEW_unidades_racks WHERE     (cd_hospital = 1)
												
												'	Set rs_racks = dbconn.execute(xsql)
												'	while Not rs_racks.EOF
												
												%>
												<a href="javascript: void(0);"  onMouseover="mostra_agendamento('block',<%=topo%>,<%=esquerda%>,'teste<%=nome_campo%>');"><input type="text" name="teste<%=nome%>" value="<%=cod_unidade%>,<%=rack%>,<%=zero(hora_ag)%>:00" size="5"></a><!-- &nbsp;unid:<%=cod_unidade%><br>rack:<%=rack%>--></td>
										<%nome_campo = nome_campo + 1
										next%>
							<%next%>
							<td><%=zero(hora_ag)%>:00</td>
						</tr>
				<%next%>
				
				<tr><td colspan="10"><%'=span_total%></td></tr>
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
	
	<div id="campo_agendamento" style="position:absolute; top:200px; left: 400px; display:none;">
		<table style="border:1px solid red; background-color:silver;">
			<tr>				
				<td>teste de posicionamento de janela: <a href="javascript: void(0);"  onclick="agendamento_abre('none')">Fecha</a></td>
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
