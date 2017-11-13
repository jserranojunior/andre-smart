<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
	<title>Untitled</title>
<%cd_unidade = request("cd_unidade")
qtd_paginas = request("qtd_paginas")
qtd_etiquetas = request("qtd_etiquetas")
mensagem = request("mensagem")
tipo_print = request("tipo_print")
	if tipo_print = 4 then
		qtd_paginas = 1
	end if
espacamento = 1
%>
<style type="text/css">
<!--
.relatorio_titulo{color: #000000;font-family: verdana;font-size:12px;font-weight:bold}
.relatorio_campos{color: #000000;font-family: verdana;font-size:7px;}
.relatorio_unidade{color: #000000;font-family: verdana;font-size:9px;font-weight:bold}
.relatorio_protocolo{color: #000000;font-family: verdana;font-size:13px;font-weight:bold}
-->
</style>
<!--onClick="window.print()-->
<SCRIPT LANGUAGE="javascript">
function Jsconfirm_print()
{
	if (confirm ("Tem certeza que deseja imprimir?"))
		{
		//document.location.href('acoes/patrimonio_acao.asp?cd_apaga='+cod1+'&cd_tipo=patrimonio');
		}
}

		
function JsDelete(cod,cod2)
{
if (confirm ("Tem certeza que deseja excluir?"))
				  {
				  document.location.href('protocolo/acoes/etiquetas_acao.asp?cod='+cod+'&acao=delete');
					}
					  }
			</SCRIPT>

</head>

<body><%'Passo 1%>
<table border="1" id="no_print" align="center" style="position:relative; top:10;">
<%if cd_unidade = "" and mensagem = "" Then '***  ***%>
	<form action="../../protocolo.asp" name="etiquetas" id="etiquetas">
	<input type="hidden" name="tipo" value="etiquetas">
	<input type="text" name="protocolo" value="<%=protocolo%>">
	<tr>
		<td colspan="2" align="center">Impressão de etiquetas para protocolo</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
	<tr>
		<td>Unidade</td>
		<td><select name="cd_unidade">
				<option value="0"></option>
				<%strsql ="Select * from TBL_unidades where cd_status <> 2  order by nm_unidade"
					Set rs_etq = dbconn.execute(strsql)%>
					<%Do while Not rs_etq.EOF%>
					<%cd_conv = rs_etq("cd_codigo")
					nm_unidade = rs_etq("nm_unidade")%>
					<%if cd_etiqueta=cd_conv Then%>
					<%ck_etq="selected"%><%else ck_etq=""%>
					<%end if%>
					<option value="<%=cd_conv%>" <%=ck_etq%>><%=nm_unidade%></option><%rs_etq.movenext
					Loop%>
			</select>
		</td>
	</tr>
<%elseif cd_unidade <> "" AND mensagem = "" then '%>	
	<form action="../../protocolo.asp" name="etiquetas" id="etiquetas">
	<input type="hidden" name="tipo" value="etiquetas">
	<input type="hidden" name="protocolo" value="<%=protocolo%>">
	<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
	<tr>
		<td colspan="2" align="center">Impressão de etiquetas para protocolo</td>
	</tr><%if cd_unidade = 0 then%><%cd_unidade = 999%><%end if%>
	<tr><td colspan="2">&nbsp;</td></tr>
	<%strsql ="Select * from TBL_protocolo_etiquetas where cd_unidade="&cd_unidade&" order by cd_num_fim desc"
			Set rs_etq = dbconn.execute(strsql)
			if not rs_etq.eof then
			cd_num_fim = rs_etq("cd_num_fim")
			else cd_num_fim = 0
			end if%>	
	<%if cd_unidade <> 999 then%>
	<tr>
		<td>Unidade</td>
		<td><%strsql ="Select * from TBL_unidades where cd_codigo = '"&cd_unidade&"'"
					Set rs_etq = dbconn.execute(strsql)
					if not rs_etq.EOF then%>
						<%=rs_etq("nm_unidade")%>						
					<%end if%>
		</td>
	</tr>
	<tr>
		<td>Última etiqueta</td>
		<td style="color:red;">&nbsp;<%=digitov(zero(cd_unidade)&"."&proto(int(cd_num_fim)))%></td>
	</tr><%cd_num_inicio = cd_num_fim + 1%>
	<tr><td>Próxima</td>
		<%num_prox = digitov(zero(cd_unidade)&"."&proto(cd_num_inicio))
		num_print_now = int(cd_num_inicio)%>
		<td style="color:green;">&nbsp;<%=num_prox%></td>
	</tr>	
	<tr>
		<td>Tipo de impressão</td><%if tipo_print = 1 OR tipo_print = "" Then%><%t_print_ck1 = "checked"%><%end if%>
		<td><input type="radio" name="tipo_print" value="1" <%=t_print_ck1%>>Somente Nº<br>
			<!--input type="radio" name="tipo_print" value="2">Somente formulário<br-->
			<%if tipo_print = 3 Then%><%t_print_ck3 = "checked"%><%end if%>
			<input type="radio" name="tipo_print" value="3" <%=t_print_ck3%>>Formulário e Nº (Completo)<br>
			<%if tipo_print = 4 Then%><%t_print_ck4 = "checked"%><%end if%>
			<input type="radio" name="tipo_print" value="4" <%=t_print_ck4%> onClick="document.etiquetas.Imprimir.click();">Folha de etiquetas</td>
	</tr>
	<%else '%>
	<input type="hidden" name="tipo_print" value="2">
	<tr>
		<td>Tipo de impressão</td>
		<td>Somente formulário </td>
	</tr>
	<tr>
		<td>Qtd. já impressas</td>
		<td style="color:red;">&nbsp;<%'=zero(cd_unidade)&"."%><%=zero(int(cd_num_fim))%></td>
	</tr><%cd_num_inicio = cd_num_fim + 1%>
	<!--tr><td>Próxima</td>
		<td style="color:green;">&nbsp;<%=zero(cd_unidade)&"."%><%=proto(cd_num_inicio)%></td>
	</tr-->
	<%end if%>
	
	<tr><td colspan="2">&nbsp;</td></tr>
	<%if tipo_print = 4 then%>
	<tr>
		<td>Quantidade de páginas</td>
		<td style="color:darkgreen;">
			<input type="hidden" name="qtd_paginas" value="1"><b>1 Folha com 65 etiquetas</b>
			<input type="hidden" name="qtd_etiquetas" value="65">
			<%qtd_etiquetas = 65%>
		</td>
	</tr>
	<%else%>
	<tr>
		<td>Quantidade de páginas</td>
		<td><select name="qtd_paginas" onBlur="document.etiquetas.Imprimir.click();">
				<%for i = 1 to 100%><%if int(qtd_paginas) = i Then%><%sel_q_pag = "selected"%><%end if%>
				<option value="<%=i%>" <%=sel_q_pag%>><%=zero(i)%></option>
				<%sel_q_pag=""%>
				<%next%>
			</select>			
		</td>
	</tr>
	<%end if%>
<%Else%>
	<%if Mensagem = 1 Then%>
	<form action="../../protocolo.asp" name="etiquetas" id="etiquetas">
	<input type="hidden" name="tipo" value="etiquetas">
	<tr><td colspan="2" style="color:green";">Impressão efetuada com sucesso!</td></tr>
	<tr><td colspan="2" align="center">Nova impressão &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </td></tr>
	<%elseif Mensagem = 2 Then%>
	<form action="../../protocolo.asp" name="etiquetas" id="etiquetas">
	<input type="hidden" name="tipo" value="etiquetas">
	<tr><td colspan="2" style="color:red";">A impressão <b>não</b> foi efetuada com sucesso!</td></tr>
	<tr><td colspan="2" align="center">Nova impressão &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </td></tr>
	<%end if%>
	
<%end if%>
	<tr>
		<td colspan="2" align="center"><input type="submit" name="Imprimir" value="Preparar impressão" width="20">&nbsp;&nbsp;&nbsp;&nbsp;<a href="protocolo.asp?tipo=etiquetas"><img src="../../imagens/undo.gif" alt="Limpar" border="0"></a><br>
	</tr>	
	</form>
</table>
<br id="no_print">
<br id="no_print">
<br id="no_print">
<table align="center" cellpadding="2" id="no_print">
	<%if cd_unidade <> "" Then
		strsql ="Select top 10 * from View_protocolo_etiquetas Where cd_unidade='"&cd_unidade&"' Order by cd_num_fim desc"
		Set rs_etiq = dbconn.execute(strsql)
	Else
		strsql ="Select top 40 * from View_protocolo_etiquetas Order by nm_unidade,cd_num_fim  desc"
		Set rs_etiq = dbconn.execute(strsql)
	End if
	%>
	<tr>
		<td colspan="100%" align="center" bgcolor="#c0c0c0"><b>Histórico de impressão de formulários</b></td>
	</tr>
	<tr>
		<td align="center"><b>Unidade</b></td>
		<td align="center"><b>Ínicio</b></td>
		<td align="center"><b>Fim</b></td>
		<td align="center"><b>Qtd.</b></td>
		<td align="center"><b>Data</b></td>
		<td align="center"><b>Usuário</b></td>
		<td align="center"><b>Status</b></td>
	</tr>
	
	<%While not rs_etiq.EOF
	
		cd_codigo = rs_etiq("cd_codigo")
		nm_unidade = rs_etiq("nm_unidade")
		cd_unidade = rs_etiq("cd_unidade")
		cd_num_inicio = rs_etiq("cd_num_inicio")
		cd_num_fim = rs_etiq("cd_num_fim")
		qtd_etq = (cd_num_fim+1) - cd_num_inicio
		dt_data = rs_etiq("dt_data")
		user_inc = rs_etiq("nm_nome")
		cd_confirm = rs_etiq("cd_confirm")
		
		
			if cor_num = 1 Then
				cor_linha = "#FFFFFF"
			else 
				cor_linha = "#dfdfdf"
			end if%>
		
		<tr bgcolor="<%=cor_linha%>">
			<td><%=nm_unidade%></td>
			<td><%=cd_num_inicio%></td>
			<td><%=cd_num_fim%></td>
			<td><%=qtd_etq%></td>
			<td><%=dt_data%></td>
			<td><%=user_inc%></td>
			<td align="center">
				<%if cd_confirm = 1 Then%>
					<img src="../../imagens/check_ok.gif" alt="Ok" border="0">					
				<%else%>
					<!--img src="../../imagens/ic_reprovado.gif" alt="Apagar" width="10" height="12" border="0"  onclick="javascript:JsDelete('<%=cd_codigo%>')"-->				
					<a href="javascript:;" onclick="window.open('protocolo/janelas/confirma_impressao.asp?cd_codigo=<%=cd_codigo%>&cd_unidade=<%=cd_unidade%>','principal','width=350, height=400'); return false;"><img src="../../imagens/ic_pendente.gif" alt="Confirmação pendente" width="10" height="12" border="0"></a>
				<%end if%>
			</td>
		</tr>
		<%cor_num = cor_num + 1
		
		if int(cor_mum) > 1 Then
			cor_num = ""
		End if
		
		rs_etiq.movenext
	wend%>
</table>



<%if cd_unidade <> "" AND qtd_paginas <> "" Then%>
<table align="center" id="no_print" style="position:relative; top:20px; left:-300px;">
	<tr>
		<td align="center" bgcolor="#ffff00" colspan="2"><b>Aviso!</b></td>
	</tr>	
	<%strsql ="Select * from TBL_unidades Where cd_codigo = "&cd_unidade&""
					Set rs_und = dbconn.execute(strsql)
					if not rs_und.EOF then
					nm_unidade = rs_und("nm_unidade")
					end if%>
	<tr><%if qtd_paginas > 1 Then 
		p1="ão" 
		p2="s" 
		Else
		p1="á"
		p2=""
		end if%>
		<td align="center"> Ser<%=p1%> impressa<%=p2%> &nbsp;<b><%=qtd_paginas%> página<%=p2%></b> do <br><b>Relatório Técnico de Cirúrgia</b><br><%if tipo_print <> 2 Then%>para o <b><%=nm_unidade%></b><%else%>em branco<%end if%>, <br>tem certeza que deseja continuar?</td>
	</tr>
	<tr>
		<td align="center">&nbsp;<br><a href="protocolo/acoes/etiquetas_acao.asp?cd_unidade=<%=cd_unidade%>&cd_num_inicio=<%=int(num_print_now)%>&qtd_paginas=<%=qtd_paginas%>&qtd_etiquetas=<%=qtd_etiquetas%>&tipo_print=<%=tipo_print%>&acao=insert&session_cd_usuario" onClick="window.print()"><b>Sim</b></a> &nbsp; &nbsp; &nbsp; <a href="protocolo.asp?tipo=etiquetas"><b>Não</b></a><p style="color:red;"><b>ATENÇÃO</b>: Após confirmar, clique em <br>preferências e certifique-se que o tamanho<br>da página está definido como <b>A4</b></p><input type="button" onclick="visualizarImpressao();" value="Visualizar Impressão" /></td>
	</tr>
</table>
<%end if

	strsql ="SELECT * FROM TBL_protocolo_etiquetas WHERE cd_unidade = '"&cd_unidade&"' ORDER BY cd_num_fim DESC"
	Set rs_proto_inicial = dbconn.execute(strsql)
		if not	rs_proto_inicial.EOF Then
			cd_num_inicio = rs_proto_inicial("cd_num_fim") + 1
		end if

protocolo = proto(int(cd_num_inicio))
' AQUI num_protocolo = cd_unidade&"."&protocolo&"-"&dv
quantidade = 0
num = 1


%>

<%do while not quantidade = int(qtd_paginas)
	if tipo_print <> 2 then
	num_protocolo = digitov(zero(cd_unidade)&"."&proto(int(protocolo)))
	end if
	'**********************************************************
	'*** DEFINE A APARENCIA EM LARGURA DE TODO O FORMULÁRIO ***
	'**********************************************************
	larg_tabs = 670
	'**************
	pad_item = round((larg_tabs/4)*0.75,0)
	pad_espaco = round((larg_tabs/4)*0.25,0)
	
	
	
	'*************************************************************%>
	 
	<%if int(qtd_paginas) = int(num) OR int(qtd_paginas) = 1 then page_break = "" else page_break = " page-break-after:always;" end if%>
	
	<%if tipo_print = 1 Then%>
		<table style="border:0px solid red; width:<%=larg_tabs%>px; left:0px;" id="ok__print">
			<tr>
				<td style="border:0px solid blue;"><img src="../../imagens/px.gif" alt="" width="55" height="1" border="0"></td>
				<td style="border:0px solid blue;">			
		<!--Cabeçalho-->
					<table style="border:0px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px;" class="relatorio_campos">
						<tr>
							<td style="width:75px; height:15px; border:0px solid purple;" rowspan="2"><!--img src="../../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"--><img src="../../imagens/px.gif" alt="" width="75" height="43" border="0"></td>
							<td style="width:<%=(round((larg_tabs-area_logotipo)/4,0))*3%>px; height:15px; border:0px solid red;" rowspan="2" align="center" class="relatorio_titulo"><!--RELATÓRIO TÉCNICO DE CIRURGIA--></td>
							<td style="width:<%=round((larg_tabs-area_logotipo)/4,0)%>px; height:15px; border:0px solid blue;" align="center" valign="bottom" class="relatorio_unidade"><%if tipo_print <> 2 then%><%=nm_unidade%><%end if%></td>
						</tr>
						<tr>
							<td style="width:179px; height:15px; border:0px solid silver;" class="relatorio_protocolo"  align="center"><%if tipo_print <> 2 then%><%=num_protocolo%><%end if%></td>
						</tr>			
					</table>
				</td>
			</tr>
		</table>
	<%elseif tipo_print = 3 OR tipo_print = 2 then%>
		<table style="border:0px solid red; width:<%=larg_tabs%>px; left:0px;" id="ok_print_">
			<tr>
				<td style="border:0px solid blue;">&nbsp;<br><br><img src="../../imagens/px.gif" alt="" width="55" height="1" border="0"><br><img src="../../imagens/centro.gif" alt="" width="7" height="26" border="0"></td>
				<td style="border:0px solid blue;">			
					<!--Cabeçalho-->
						<table style="border:0px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px;" class="relatorio_campos">
							<tr>
								<td style="width:75px; height:15px; border:0px solid purple;" rowspan="2"><img src="../../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"><%area_logotipo="75"%></td>
								<td style="width:<%=(round((larg_tabs-area_logotipo)/4,0))*3%>px; height:15px; border:0px solid red;" rowspan="2" align="center" class="relatorio_titulo">RELATÓRIO TÉCNICO DE CIRURGIA</td>
								<td style="width:<%=round((larg_tabs-area_logotipo)/4,0)%>px; height:15px; border:0px solid blue;" align="right" valign="bottom" class="relatorio_unidade"><%if tipo_print <> 2 then%><%=nm_unidade%><%end if%></td>
							</tr>
							<tr>
								<td style="width:179px; height:15px; border:0px solid silver;" class="relatorio_protocolo"  align="center"><%if tipo_print <> 2 then%><%=num_protocolo%><%end if%></td>
							</tr>			
						</table>	
					<br><!--Etiqueta e data-->
						<table style="border:0px solid black; border-collapse:collapse; width:<%=round(larg_tabs,0)%>px;" class="relatorio_campos">
							<tr>
								<td style="border:1px; width:358px; height:80px;" align="left">
									<table style="border:2px dotted silver; border-collapse:collapse; width:350px;" class="relatorio_campos">
										<tr><td style="width:350px; height:80px; border:1px solid silver; color:silver;" align="center">ETIQUETA<%'="<br>"&pad_item%><%'=<br>"&pad_espaco%></td></tr>					
									</table>
								</td>
								<td style="border:1px width:<%=round(larg_tabs/5,0)%>px; height:80px;" align="left">&nbsp;</td>
								<td style="border:1px width:<%=round(larg_tabs/4,0)%>px; height:15px;" align="right">
									<table style="width:<%=round(larg_tabs/4,0)%>px; border:2px solid silver; border-collapse:collapse;" class="relatorio_campos">
										<tr><td style="width:<%=round(larg_tabs/4,0)%>px; height:20px; border:1px solid silver;">&nbsp;Data</td></tr>
										<tr><td style="width:<%=round(larg_tabs/4,0)%>px; height:20px; border:1px solid silver;">&nbsp;Hora agendada</td></tr>
										<tr><td style="width:<%=round(larg_tabs/4,0)%>px; height:20px; border:1px solid silver;">&nbsp;Hora início</td></tr>
										<tr><td style="width:<%=round(larg_tabs/4,0)%>px; height:20px; border:1px solid silver;">&nbsp;Hora término</td></tr>
									</table>
								</td>
							</tr>
						</table>	
					<br><!--Cirurgião-->
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px;" class="relatorio_campos">
							<tr>
								<td style="width:<%=(round(larg_tabs/4,0))*3%>px; height:18px; border:1px solid silver;">&nbsp;Cirurgião</td>
								<td style="width:<%=round(larg_tabs/4,0)%>px; height:18px; border:1px solid silver;">&nbsp;CRM</td>			
							</tr>
							<tr>
								<td colspan="2" style="width:<%=larg_tabs%>px; height:12px; border:1px solid silver;">&nbsp;Tipo de rack: ( &nbsp; ) Geral&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Ginecologia&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Histeroscopia&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Neurologia&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Ortopedia&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Otorrino&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Torax&nbsp;&nbsp;&nbsp;( &nbsp;&nbsp; ) Urologia</td>
							</tr>		
						</table>
					<br><!--Cirurgia-->
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative; " class="relatorio_campos">
							<tr>
								<td colspan="2" style="width:<%=larg_tabs%>px; height:12px; border:1px solid silver;">&nbsp;Eletiva ( )&nbsp;&nbsp;&nbsp;&nbsp;Pós-Agendada ( )&nbsp;&nbsp;&nbsp;&nbsp;Urgência()</td>
							</tr>
							<tr>
								<td style="width:<%=(round(larg_tabs/4,0))*3%>px; height:12px; border:1px solid silver;">&nbsp;Cirurgia(s) realizada(s)</td>
								<td style="width:<%=round(larg_tabs/4,0)%>px; height:12px; border:1px solid silver;">&nbsp;Cód. AMB</td>
							</tr>
							<tr>
								<td style="width:<%=(round(larg_tabs/4,0))*3%>px; height:18px; border:1px solid silver;">&nbsp;1</td>
								<td style="width:<%=round(larg_tabs/4,0)%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr>
								<td style="width:<%=(round(larg_tabs/4,0))*3%>px; height:18px; border:1px solid silver;">&nbsp;2</td>
								<td style="width:<%=round(larg_tabs/4)%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr>
								<td style="width:<%=(round(larg_tabs/4,0))*3%>px; height:18px; border:1px solid silver;">&nbsp;3</td>
								<td style="width:<%=round(larg_tabs/4)%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr>
								<td style="width:<%=(round(larg_tabs/4,0))*3%>px; height:18px; border:1px solid silver;">&nbsp;4</td>
								<td style="width:<%=round(larg_tabs/4)%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>
						</table>
					<br><!--Esterilização-->
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative;" class="relatorio_campos">
							<tr>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:30px; border:1px solid silver;" valign="top" rowspan="3">&nbsp;Armário<br><br>&nbsp;Limpeza com alcool a 70%<br><br>&nbsp;SIM ( &nbsp; ) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Não ( &nbsp; )</td>			
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:30px; border:1px solid silver;" valign="top">&nbsp;Sterrad<br><br>&nbsp;Data</td>			
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:30px; border:1px solid silver;" valign="top">&nbsp;Autoclave<br><br>&nbsp;Data</td>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:30px; border:1px solid silver;" valign="top">&nbsp;Statin<br><br>&nbsp;Data</td>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:30px; border:1px solid silver;" valign="top" rowspan="3">&nbsp;Anti-sepsia da pele do paciente<br><br> &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SIM ( )&nbsp;&nbsp;Não ( )<br><br>&nbsp;PVPI&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;( )<br><br>&nbsp;Clorexidina &nbsp; &nbsp;( )</td>
							</tr>
							<tr>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:18px; border:1px solid silver;" valign="top">&nbsp;Validade</td>			
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:18px; border:1px solid silver;" valign="top">&nbsp;Validade</td>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:18px; border:1px solid silver;" valign="top">&nbsp;Validade</td>
							</tr>
							<tr>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:18px; border:1px solid silver;">&nbsp;</td>			
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:18px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=round(larg_tabs/5,0)%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>
						</table>
					<br><!--Materiais-->
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative;" class="relatorio_campos">
							<tr><!--linha 1.0-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Equipamentos</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">Rack Num.</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Ótica</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Bandeja (Qtd. de peças)</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
						<!--/table>
						<br>
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative;" class="relatorio_campos"-->	
							<tr>
								<td style="height:15px; border:1px solid silver;" colspan="8">&nbsp;<b>Laparoscopia</b></td>			
							</tr>
							<tr><!--linha 1.1-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Afastador de fígado</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Mandril 10mm</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça Pozzi</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Vedantes 5 mm - externos</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.2-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Agulha de verrez</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Mandril 5mm</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça saca mioma</td>			
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.3-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Apalpador</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Manipulador uterino</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Porta agulha</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.4-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Cabo bipolar</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça atraumática de trompa</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Redutor metálico</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.5-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Cabo monopolar</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça babcock</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Tesoura</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.6-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Cânula punção</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça biópsia</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Trocater 05 mm</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.7-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Cauterizador e irrigador</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça endoclinch</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Trocater 10 mm</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.8-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Clipador</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça grasper com dente</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Vedantes 10 mm - internos</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.9-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Conector monopolar</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça grasper curva</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Vedantes 10 mm - externos</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 1.10-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Contra-porta agulha</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça grasper reta</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Vedantes 5 mm - internos</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
						<!--/table>
						<br>
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative;" class="relatorio_campos"-->	
							<tr>
								<td style="height:15px; border:1px solid silver;" colspan="8">&nbsp;<b>Urologia / Histeroscopia</b></td>			
							</tr>
							<tr><!--linha 2.1-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Cabo de alta freqüência</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Camisa Diagnostica</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça flexível p/ corpo estranho</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 2.2-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Camisa 20</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Elemento de trabalho – 1 pino</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Ponte 1 via</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 2.3-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Camisa 21</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Elemento de trabalho – 2 pinos</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Ponte 2 vias</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 2.4-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Camisa 22</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Evacuador de Ellick</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 2.5-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Camisa 24</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Faca sachs</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr><!--linha 2.6-->
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Camisa 26</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;Pinça flexível p/ biópsia</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:1px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:1px solid silver;">&nbsp;</td>
							</tr>
						</table>
						<table style="border:0px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative;" class="relatorio_campos">
							<tr style="border:0px;"><!--linha 2.6-->
								<td style="width:<%=pad_item%>px; height:13px; border:0px solid silver;" colspan="6">&nbsp;</td>
								<!--td style="width:<%=pad_espaco%>px; height:13px; border:0px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:0px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:0px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_item%>px; height:13px; border:0px solid silver;">&nbsp;</td>
								<td style="width:<%=pad_espaco%>px; height:13px; border:0px solid silver;">&nbsp;</td-->
								<td style="width:<%=pad_item-2%>px; height:13px; border:0px solid silver;" align="right">&nbsp;<b>TOTAL</b>&nbsp;</td>
								<td style="width:<%=pad_espaco-2%>px; height:13px; border:2px solid silver;">&nbsp;</td>
							</tr>
						</table>
					<br><!--Equipamentos e instrumentos externos-->
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative; " class="relatorio_campos">
							<tr>
								<td colspan="2" style="width:<%=larg_tabs%>px; height:13px; border:1px solid silver;">&nbsp;Equipamentos e Instrumentos Externos</td>
							</tr>
							<tr>
								<td colspan="2" style="width:<%=larg_tabs%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="2" style="width:<%=larg_tabs%>px; height:18px; border:1px solid silver;">&nbsp;</td>
							</tr>		
						</table>
					<!--Observações-->
						<div style="position:relative; position:relative; top:5;" class="relatorio_campos">Observações / Intercorrências</div>
						<hr width="<%=larg_tabs%>" align="left" style="position:relative; position:relative;" color="silver" noshade size="1"><br>
						<hr width="<%=larg_tabs%>" align="left" style="position:relative; position:relative;" color="silver" noshade size="1"><br>
						<hr width="<%=larg_tabs%>" align="left" style="position:relative; position:relative;" color="silver" noshade size="1"><br>
						<hr width="<%=larg_tabs%>" align="left" style="position:relative; position:relative;" color="silver" noshade size="1"><br>
						<hr width="<%=larg_tabs%>" align="left" style="position:relative; position:relative;" color="silver" noshade size="1"><br>
							<!--hr width="<%=larg_tabs%>" align="left" style="position:relative; position:relative;" color="silver" noshade size="1"><br-->
						<%larg_campo1 = (pad_item*2)+pad_espaco
						larg_campo2 = (pad_item*2)+(pad_espaco*3)%>
					<!-- Carimbos e assinaturas-->
						<table style="border:2px solid silver; border-collapse:collapse; width:<%=larg_tabs%>px; position:relative;" class="relatorio_campos">
							<tr>
								<td style="width:<%=larg_campo1%>px; height:40px; border:1px solid silver;" valign="top">&nbsp;Circulante de Sala</td>
								<td style="width:<%=larg_campo2%>px; height:40px; border:1px solid silver;" valign="top">&nbsp;Gravação DVD &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; SIM ( )  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Não ( )<br><br>&nbsp;Entregue para</td>						
							</tr>
							<tr>
								<td style="width:<%=larg_campo1%>px; height:40px; border:1px solid silver;" valign="top">&nbsp;Enfermagem Técnica CMI</td>
								<td style="width:<%=larg_campo2%>px; height:80px; border:1px solid silver;" rowspan="2" valign="top">&nbsp;Observações
									<p style="position: relative; bottom:0;" align="right"><br><br><br><br><br>
									Médico / Enfermeira&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</p></td>						
							</tr>
							<tr>
								<td style="width:<%=larg_campo1%>px; height:45px; border:1px solid silver;" valign="top">&nbsp;Coordenação Técnica &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; / &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; retirado por</td>			
							</tr>
						</table><!--img src="../../imagens/blackdot.gif" alt="" width="200" height="1" border="0"-->
					<br>
						<table style="border:2px dotted silver; border-collapse:collapse; width:350px; position:relative;" class="relatorio_campos">
							<tr>
								<td style="height:60px; border:1px solid silver; color:silver;" align="center">Etiqueta de esterilização</td>
							</tr>
						</table>
		</td>
			</tr>
		</table>
		<br><br>
	<%Elseif tipo_print = 4 then%>
	 <% '**************************************************************
		'*** DEFINE A APARENCIA DA FOLHA DE ETIQUETA PIMACO - A4251 ***
		'**************************************************************
		coluna = "1"
		linha = "1"
		numero = "0"
		msup = "22" 'Borda superior da etiqueta
		'msup = "105"
		mesq =  "3" 'Borda esquerda da etiqueta
		 
		msup = msup + 6 'posição do texto (superior)
		mesq = mesq 'posição do texto (esquerda). O texto já está centralizado
		 
		qtd_etiquetas = 13
		%>
		<table style="border:1px solid white; border-collapse:collapse; position:relative; width:760px; left:2px; top:4;" id="ok__print">
			<tr>
					<td style="border:0px solid blue; height:20px; width:<%=larg_etiq%>;" align="center"></td>
					<td style="border:0px solid blue; height:px; width:<%=espac_etiq%>;"></td>
					<td style="border:0px solid blue; height:px; width:<%=larg_etiq%>;" align="center"></td>
					<td style="border:0px solid blue; height:px; width:<%=espac_etiq%>;"></td>
					<td style="border:0px solid blue; height:px; width:<%=larg_etiq%>;" align="center"></td>
					<td style="border:0px solid blue; height:px; width:<%=espac_etiq%>;"></td>
					<td style="border:0px solid blue; height:px; width:<%=larg_etiq%>;" align="center"></td>
					<td style="border:0px solid blue; height:px; width:<%=espac_etiq%>;"></td>
					<td style="border:0px solid blue; height:px; width:<%=larg_etiq%>;" align="center"></td>
				</tr>
		<%while not numero = int(qtd_etiquetas)%>			
			<%larg_etiq = 145
			Altura_etiq = 80
			espac_etiq = 5%>
				<tr>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px; width:<%=larg_etiq%>;" align="center"><br>
						<font color="#0000ff" face="Arial" size="1"><%=nm_unidade%></font><br><%'=int(qtd_paginas)%>
						&nbsp;<br>
						<font color="#0000ff" face="Arial" size="2"><b><%=digitov(zero(cd_unidade)&"."&proto(protocolo))%></b></font></td>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px; width:<%=espac_etiq%>;"></td>
					<%protocolo = protocolo + 1%>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px;" align="center"><br>
						<font color="#0000ff" face="Arial" size="1">&nbsp;<%=nm_unidade%></font><br><%'=int(qtd_paginas)%>
						&nbsp;<br>
						<font color="#0000ff" face="Arial" size="2">&nbsp;<b><%=digitov(zero(cd_unidade)&"."&proto(protocolo))%></b></font></td>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px; width:<%=espac_etiq%>;"></td>
					<%protocolo = protocolo + 1%>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px;" align="center"><br>
						<font color="#0000ff" face="Arial" size="1">&nbsp;<%=nm_unidade%></font><br><%'=int(qtd_paginas)%>
						&nbsp;<br>
						<font color="#0000ff" face="Arial" size="2">&nbsp;<b><%=digitov(zero(cd_unidade)&"."&proto(protocolo))%></b></font></td>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px; width:<%=espac_etiq%>;"></td>
					<%protocolo = protocolo + 1%>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px;" align="center"><br>
						<font color="#0000ff" face="Arial" size="1">&nbsp;<%=nm_unidade%></font><br><%'=int(qtd_paginas)%>
						&nbsp;<br>
						<font color="#0000ff" face="Arial" size="2">&nbsp;<b><%=digitov(zero(cd_unidade)&"."&proto(protocolo))%></b></font></td>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px; width:<%=espac_etiq%>;"></td>
					<%protocolo = protocolo + 1%>
					<td style="border:0px solid blue; height:<%=altura_etiq%>px;" align="center"><br>
						<font color="#0000ff" face="Arial" size="1">&nbsp;<%=nm_unidade%></font><br><%'=int(qtd_paginas)%>
						&nbsp;<br>
						<font color="#0000ff" face="Arial" size="2">&nbsp;<b><%=digitov(zero(cd_unidade)&"."&proto(protocolo))%></b></font></td>
				</tr>
				
			
			
			<%	if coluna = 5 then
					coluna = 0
					linha = linha + 1
					mesq = 3
				else
					msup = msup '+ 100
					mesq = mesq + 154
				end if
				
				coluna = coluna + 1
				numero = numero + 1
				protocolo = protocolo + 1
			
		wend%>
		</table>
	<%end if%>
<div style="<%=page_break%>" align="center"></div>	
<%
numero = 1
protocolo = protocolo + 1
quantidade = quantidade + 1
num = num + 1
espacamento = espacamento + 1
loop
%>

</body>
</html>
