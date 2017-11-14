<!--#include file="../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%no_patrimonio = request("no_patrimonio")
'if cd_user = 46 then response.write("ajax/ajax_os_pat_NS_mostra.asp")
no_serie = request("no_serie")

nm_tipo = request("tipo_pat")

wdt_nome_item = 75
wdt_campo_item = 200%>
	<table>
<%if len(no_patrimonio) > 2 OR len(no_serie) > 2 Then

		if isNumeric(no_patrimonio) = true Then
			xsql = "SELECT count(cd_codigo) as conta_reg FROM View_patrimonio_lista Where no_patrimonio = '"&no_patrimonio&"' and nm_tipo = '"&nm_tipo&"'"
			Set rs = dbconn.execute(xsql)
			conta_reg = rs("conta_reg")
			
			xsql = "SELECT * FROM View_patrimonio_lista Where no_patrimonio = '"&no_patrimonio&"' and nm_tipo = '"&nm_tipo&"'"
			Set rs = dbconn.execute(xsql)
		end if
		
		if no_serie <> "" Then
			xsql = "SELECT count(cd_codigo) as conta_reg FROM View_patrimonio_lista Where cd_ns like '"&no_serie&"%' and nm_tipo = '"&nm_tipo&"'"
			Set rs = dbconn.execute(xsql)
			conta_reg = rs("conta_reg")
			
			xsql = "SELECT * FROM View_patrimonio_lista Where cd_ns like '"&no_serie&"%' and nm_tipo = '"&nm_tipo&"'"
			Set rs = dbconn.execute(xsql)
		end if
		
		While not rs.EOF 
			cd_pat = rs("cd_codigo")
			num_patrimonio = rs("no_patrimonio")
			cd_equipamento = rs("cd_equipamento")
			nm_equipamento = rs("nm_equipamento")
			cd_marca = rs("cd_marca")
			nm_marca = rs("nm_marca")
			cd_especialidade = rs("cd_especialidade")
			nm_especialidade = rs("nm_especialidade")
			cd_ns = rs("cd_ns")
			if no_patrimonio <> "" Then
				cor_num_pat = "red"
				cor_num_serie = "black"
			elseif no_serie <> "" Then
				cor_num_pat = "black"
				cor_num_serie = "red"
			end if
		%>
				<tr>
					<td rowspan="3" align="center" valign="middle" style="width:20px;">
						<%if int(conta_reg) > 1 then%><input type="radio" name="seleciona_patrimonio" value="" onClick="alimenta_info_pat(<%="'"&cd_pat&"','"&no_patrimonio&"','"&cd_equipamento&"','"&cd_marca&"','"&cd_especialidade&"','"&cd_ns&"','1'"%>);"><%end if%>
					</td>
					<td style="width:<%=wdt_nome_item%>px; color:<%=cor_num_pat%>;"><b>Pat:</b></td>
					<td style="width:<%=wdt_campo_item%>px; color:#339900;"><b><%=num_patrimonio%></b></td>					
				</tr>
			  	<tr>
					<td style="width:<%=wdt_nome_item%>px;"><b>Item:</b></td>
					<td style="width:<%=wdt_campo_item%>px;"><%=nm_equipamento%></td>
					<td style="width:<%=wdt_nome_item%>px;"><b>Especialidade:</b></td>
					<td style="width:50px;"><%=nm_especialidade%></td>			
				</tr>
				<tr>
					<td style="width:<%=wdt_nome_item%>px;"><b>Marca:</b></td>
					<td style="width:<%=wdt_campo_item%>px;"><%=nm_marca%></td>
					<td style="width:<%=wdt_nome_item%>px; color:<%=cor_num_serie%>;"><b>N/S:</b></td>
					<td style="width:50px;"><%=cd_ns%></td>
				</tr>
				<tr><td colspan="5"><img src="../../imagens/blackdot.gif" alt="" width="450" height="1" border="0"></td></tr>
		
		<%
		cor_num_pat = "black"
		cor_num_serie = "black"
		%>				
				
			<%reg_conta = reg_conta + 1%>
			<tr>
				<td colspan="4">
					<!--span id="divPat_reg"-->
						<!-- Quando a imagem abaixo for carregada, envia os dados do equip. para os campos necessários -->
						<%if int(conta_reg) <= 1 then%>
							<img src="../../imagens/px.gif" alt="" width="1" height="1" border="0" id="myImage" onload="alimenta_info_pat(<%="'"&cd_pat&"','"&no_patrimonio&"','"&cd_equipamento&"','"&cd_marca&"','"&cd_especialidade&"','"&cd_ns&"','1'"%>);">
						<%else%>
							<img src="../../imagens/px.gif" alt="" width="1" height="1" border="0" id="myImage" onload="alimenta_info_pat('','','','','','','');">
						<%end if%>						
					<!--/span-->
				</td>
			</tr>
			<%rs.movenext
			wend
			
				rs.close
			 	set rs = nothing
			  if reg_conta = 0 then%>
				
			   <tr><td>O patrimômio ou numero de série informado não existe.</td></tr>
			  <%end if%>
		</table>
	<%'else%>
		<!--span style="color:red;">Digite apenas números.<%=cd_patrimonio%></span-->
	<%'end if%>	
<%'Else%>
	<!--table>
		<tr>
			<td style="width:380px;">*** Nenhum registro encontrado. ***</td>
		</tr>
	</table-->
<%end if%>
<script language="JavaScript">

</script>