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
		<table>
			<tr>
				<td>Listagem de Movimentção de patrimônio</td>
			</tr>
			<tr><td>&nbsp;</td></tr>
			<tr>
				<td>Patrimônio:</td><td><%=no_patrimonio%></td>
				<td>N/S:</td><td><%=cd_ns%></td>
			</tr>
			<tr>
				<td>Marca:</td><td><%=cd_marca%></td>
				<td>Aquisição:</td><td><%=dt_data%></td>
			</tr>
		</table>


</body>
</html>
