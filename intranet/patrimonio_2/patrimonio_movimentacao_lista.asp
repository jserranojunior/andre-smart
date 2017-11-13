<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<head>
	<title>Untitled</title>
</head>
<%cd_pat_codigo = request("cd_pat_codigo")%>
<body>
	<%strsql ="SELECT * FROM View_patrimonio_lista WHERE cd_codigo = "&cd_pat_codigo&""
	  	Set rs_patrimonio = dbconn.execute(strsql)
		
		IF not rs_patrimonio.EOF THEN
			cd_pat_codigo = rs_patrimonio("cd_codigo")
			no_patrimonio = rs_patrimonio("no_patrimonio")
			cd_equipamento = rs_patrimonio("cd_equipamento")
			nm_equipamento = rs_patrimonio("nm_equipamento_novo")
			cd_marca = rs_patrimonio("cd_marca")
			nm_marca = rs_patrimonio("nm_marca")
			cd_ns = rs_patrimonio("cd_ns")
			cd_pn = rs_patrimonio("cd_pn")
			nm_tipo = rs_patrimonio("nm_tipo")
			cd_especialidade = rs_patrimonio("cd_especialidade")
			cd_unidade = rs_patrimonio("cd_unidade")
			nm_unidade = rs_patrimonio("nm_unidade")
			nm_sigla = rs_patrimonio("nm_sigla")
			
			cd_rack = rs_patrimonio("cd_rack")
			nm_rack = rs_patrimonio("nm_rack")
			
			cd_unidade_comodato = rs_patrimonio("cd_unidade_comodato")
			cd_rack_comodato = rs_patrimonio("cd_rack_comodato")
			nm_unidade_comodato = rs_patrimonio("nm_unidade_comodato")
			nm_rack_comodato = rs_patrimonio("nm_rack_comodato")
			
			num_hospital =  rs_patrimonio("num_hospital")
			num_hospital = rs_patrimonio("num_hospital")
			cd_categoria = rs_patrimonio("cd_categoria")
			dt_data = rs_patrimonio("dt_data")
			
			cd_preventiva = rs_patrimonio("cd_preventiva")
			dt_periodicidade_prev = rs_patrimonio("dt_periodicidade_prev")
			
			cd_calibracao = rs_patrimonio("cd_calibracao")
			dt_periodicidade_cal = rs_patrimonio("dt_periodicidade_cal")
			
			cd_seg_elet = rs_patrimonio("cd_seg_elet")
			dt_periodicidade_elet = rs_patrimonio("dt_periodicidade_elet")
			
			nao_constar = rs_patrimonio("nao_constar")
			cd_status = rs_patrimonio("cd_status")
				if cd_status = 2 then
					cor_tabela = "#f0f0f0"
					nm_status = "<span style='color:red;'> ( &nbsp; (&nbsp; ( D E S C A R T A D O ) &nbsp;) &nbsp; ) </span>"
				else
					cor_tabela = "#f0f0f0"
					nm_status = ""
				end if
			
			condicao = " Where dt_descarte is null"
			ordem = " "&ordem&" "
			'ordem = " cd_codigo desc"
				'strsql_espec ="SELECT * FROM TBL_ESPEC Where cd_codigo='"&cd_especialidade&"'"
				strsql_espec ="SELECT * FROM TBL_ESPECialidades Where cd_codigo='"&cd_especialidade&"'"
			  	Set rs_espec = dbconn.execute(strsql_espec)
					if not rs_espec.EOF then
						nm_especialidade = rs_espec("nm_especialidade")
					end if
			'rs_espec.movenext
		End if%>
		<table align="center">
			<tr>
				<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="125" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="125" height="1" border="0"></td>				
			</tr>
			<tr>
				<td align="center" colspan="4" style="background-color:gray; color:white;"><b>Listagem de Movimentção de patrimônio</b></td>
			</tr>
			<tr>
				<td>Patrimônio:</td><td><b><%=nm_tipo&no_patrimonio%></b></td>
				<td>N/S:</td><td><b><%=cd_ns%></b></td>
			</tr>
			<tr>
				<td>Item:</td><td colspan="4"><b><%=nm_equipamento%></b></td>
			</tr>
			<tr>
				<td>Marca:</td><td><b><%=nm_marca%></b></td>
				<td>Aquisição:</td><td><b><%=zero(day(dt_data))&"/"&zero(month(dt_data))&"/"&year(dt_data)%></b></td>
			</tr>
			<tr><td colspan="4" style="background-color:gray;"></td></tr>
		</table>
		<table align="center">	
			<%strsql = "SELECT * FROM VIEW_patrimonio_movimentacao_lista WHERE (cd_patrimonio="&cd_pat_codigo&") and id_movimento = 1 and cd_status = 1"
			Set rs = dbconn.execute(strsql)
			
			if not rs.EOF then
				cd_unidade_origem = rs("cd_unidade_origem")
				cd_unidade_destino = rs("cd_unidade_destino")%>
				<tr style="background-color:silver;">
						<td>&nbsp;</td>
						<td>Origem</td>				
						<td>Destino</td>
						<td>Ida</td>
						<td>Volta</td>
						<td>Tipo</td>				
					</tr>
				<%'strsql = "SELECT * FROM VIEW_patrimonio_movimentacao_lista WHERE (cd_patrimonio="&cd_pat_codigo&") and id_movimento = 1 and cd_status = 1"
				strsql = "SELECT * FROM TBL_empresa WHERE cd_codigo = "&cd_unidade_origem&""
				Set rs_empresa = dbconn.execute(strsql)%>
				
				<tr style="background-color:#e9e9e9;">
					<td>&nbsp;1</td>
					<td><b><%=rs_empresa("nm_empresa_abrv")%></b></td>				
					<td><%=rs("nm_sigla_destino")%></td>
					<td><%=zero(day(rs("dt_ida")))&"/"&zero(month(rs("dt_ida")))&"/"&year(rs("dt_ida"))%></td>
					<td><%%></td>
					<td><%=rs("nm_situacao")%></td>				
				</tr>
			<%cabeca = 1
			end if%>
			
			<%n_linha = 1
			strsql = "SELECT * FROM VIEW_patrimonio_movimentacao_lista WHERE (cd_patrimonio="&cd_pat_codigo&") and id_movimento > 1 and cd_status = 1"
			Set rs = dbconn.execute(strsql)
			
			while not rs.EOF
				cd_unidade_origem = rs("cd_unidade_origem")
				nm_unidade_origem = rs("nm_sigla_origem")
				cd_unidade_destino = rs("cd_unidade_destino")
				nm_unidade_destino = rs("nm_sigla_destino")
				dt_ida = rs("dt_ida")
					dt_ida = zero(day(dt_ida))&"/"&zero(month(dt_ida))&"/"&year(dt_ida)
				dt_retorno = rs("dt_retorno")
					dt_retorno = zero(day(dt_retorno))&"/"&zero(month(dt_retorno))&"/"&year(dt_retorno)
				
				cd_tipo_movimentacao = rs("cd_tipo_movimentacao")
				'	if cd_tipo_movimentacao = 1 then
				'		nm_tipo_movimentacao = "Descarte"
				'	elseif cd_tipo_movimentacao = 2 then
				'		nm_tipo_movimentacao = "Manutenção"
				'	elseif cd_tipo_movimentacao = 3 then
				'		nm_tipo_movimentacao = "Aquisição"
				'	elseif cd_tipo_movimentacao = 4 then
				'		nm_tipo_movimentacao = "Empréstimo"
				'	elseif cd_tipo_movimentacao = 5 then
				'		nm_tipo_movimentacao = "Comodato"
				'	end if
				nm_tipo_movimentacao = rs("nm_situacao")
				n_linha = n_linha + 1
				
				if cabeca = 0 then%>
					<tr>
						<td>&nbsp;</td>
						<td>Origem</td>				
						<td>Destino</td>
						<td>Ida</td>
						<td>Volta</td>
						<td>Tipo</td>				
					</tr>
				<%cabeca = 1
				end if%>
				<tr>
					<td>&nbsp;<%=n_linha%></td>
					<td><%=nm_unidade_origem%></td>				
					<td><%=nm_unidade_destino%></td>
					<td><%=dt_ida%></td>
					<td><%=dt_retorno%></td>
					<td><%=nm_tipo_movimentacao%></td>				
				</tr>
			<%if num_linha = 40 then
				cabeca = 0
			end if
			rs.movenext
			wend%>
			<tr><td colspan="6" style="background-color:gray;"></td></tr>
			<tr>
				<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="90" height="1" border="0"></td>
				<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
			</tr>
		</table>


</body>
</html>
