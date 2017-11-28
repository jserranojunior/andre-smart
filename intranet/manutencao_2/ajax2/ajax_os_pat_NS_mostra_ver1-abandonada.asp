<!--#include file="../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%no_patrimonio = request("no_patrimonio")
nm_tipo = request("tipo_pat")

'pos_pat = instr(1,cd_patrimonio,":",1)
'nm_tipo = mid(cd_patrimonio,1,1)
'no_patrimonio = mid(cd_patrimonio,3,len(cd_patrimonio))
%>
	<table>
<%if len(no_patrimonio) > 2 Then

		if isNumeric(no_patrimonio) = true Then
			xsql = "SELECT count(no_patrimonio) as conta_pat FROM View_patrimonio_lista Where no_patrimonio = '"&no_patrimonio&"' and nm_tipo = '"&nm_tipo&"'"
			Set rs = dbconn.execute(xsql)
			conta_pat = rs("conta_pat")
		else
			conta_pat = 0
		end if
		
		xsql2 = "SELECT count(cd_codigo) as conta_ns FROM View_patrimonio_lista Where cd_ns Like '"&no_patrimonio&"%' and nm_tipo = '"&nm_tipo&"'"
		Set rs2 = dbconn.execute(xsql2)
		conta_ns = rs2("conta_ns")
		
		response.write(conta_pat&"-"&conta_ns&"<br><br>")
		
		
		soma_pat_ns = conta_pat + conta_ns
		
		if conta_pat > 0 then
			xsql = "SELECT * FROM View_patrimonio_lista Where no_patrimonio = '"&no_patrimonio&"' and nm_tipo = '"&nm_tipo&"'"
			Set rs = dbconn.execute(xsql)
		end if
		
		if conta_ns > 0 then
			xsql2 = "SELECT * FROM View_patrimonio_lista Where cd_ns like '"&no_patrimonio&"%' and nm_tipo = '"&nm_tipo&"'"
			Set rs2 = dbconn.execute(xsql2)
		end if
		
		x_pat 	= 1
		x_ns	= 1
		n		= 1
		'response.write("<br>("&soma_pat_ns&")<br>")
		
		for i = 1 to soma_pat_ns
			
			
			
			if conta_pat > 0 AND x_pat <= conta_pat then
				
				'response.write(rs("cd_codigo")&" rs "&x_pat&"<br>")
					cd_pat = rs("cd_codigo")
					num_patrimonio = rs("no_patrimonio")
					cd_equipamento = rs("cd_equipamento")
					nm_equipamento = rs("nm_equipamento")
					cd_marca = rs("cd_marca")
					nm_marca = rs("nm_marca")
					cd_especialidade = rs("cd_especialidade")
					nm_especialidade = rs("nm_especialidade")
					cd_ns = rs("cd_ns")
						cor_num_pat = "red"
						cor_num_serie = "black"
				rs.movenext
				x_pat = x_pat + 1
			'end if
			
			Elseif conta_ns > 0 and x_ns <= conta_ns then
				
				'response.write(rs2("cd_codigo")&" rs2 "&x_pat&" i="&i&"<br>bbb")
					cd_pat = rs2("cd_codigo")
					num_patrimonio = rs2("no_patrimonio")
					cd_equipamento = rs2("cd_equipamento")
					nm_equipamento = rs2("nm_equipamento")
					cd_marca = rs2("cd_marca")
					nm_marca = rs2("nm_marca")
					cd_especialidade = rs2("cd_especialidade")
					nm_especialidade = rs2("nm_especialidade")
					cd_ns = rs2("cd_ns")
					cor_num_pat = "black"
					cor_num_serie = "red"
				rs2.movenext
			x_ns = x_ns + 1
			
			end if
			
			
		
		'if reg_conta >= 1 then%>
				<!--tr><td colspan="5" style="background-color:gray;"></td></tr-->
				<%'end if%>
				<tr>
					<td rowspan="3" align="center" valign="middle" style="width:20px;">
						<%if int(conta_pat) > 1 OR int(conta_ns) > 1 then%><input type="radio" name="seleciona_patrimonio" value="" onClick="alimenta_info_pat(<%="'"&cd_pat&"','"&no_patrimonio&"','"&cd_equipamento&"','"&cd_marca&"','"&cd_especialidade&"','"&cd_ns&"','1'"%>);"><%end if%>
					</td>
					<td style="width:50px; color:<%=cor_num_pat%>;"><b>Pat:</b></td>
					<td style="width:200px; color:#339900;"><b><%=num_patrimonio%></b></td>
					<td><%=n%>.<%=cd_pat%></td>
				</tr>
			  	<tr>
					<td style="width:50px;"><b>Item:</b></td>
					<td style="width:200px;"><%=nm_equipamento%></td>
					<td style="width:50px;"><b>Especialidade:</b></td>
					<td style="width:50px;"><%=nm_especialidade%></td>			
				</tr>
				<tr>
					<td style="width:50px;"><b>Marca:</b></td>
					<td style="width:200px;"><%=nm_marca%></td>
					<td style="width:50px; color:<%=cor_num_serie%>;"><b>N/S:</b></td>
					<td style="width:50px;"><%=cd_ns%></td>
				</tr>
				<tr><td colspan="4">...</td></tr>
		
		
		<%
		
		cor_num_pat = "black"
		cor_num_serie = "black"
		
		'cd_pat = ""
		'num_patrimonio = ""
		'cd_equipamento = ""
		'nm_equipamento = ""
		'cd_marca = ""
		'nm_marca = ""
		'cd_especialidade = ""
		'nm_especialidade = ""
		'cd_ns = ""
		'cor_num_pat = ""
		'cor_num_serie = ""
		'if n = conta_pat then n = 0
		'n = n + 1
		
		next
		%>

		
			
				
									
				
			<%reg_conta = reg_conta + 1%>
			<tr>
				<td colspan="4"><%'=int(conta_pat)&" . "&int(conta_ns)%>
					<!--span id="divPat_reg"-->
						<!-- Quando a imagem abaixo for carregada, envia os dados do equip. para os campos necessários -->
						<%if int(conta_pat) <= 1 OR int(conta_ns) <= 1 then%>
							<img src="../../imagens/blackdot.gif" alt="" width="150" height="1" border="0" id="myImage" onload="alimenta_info_pat(<%'="'"&cd_pat&"','"&no_patrimonio&"','"&cd_equipamento&"','"&cd_marca&"','"&cd_especialidade&"','"&cd_ns&"','1'"%>);"><%="'"&cd_pat&"','"&no_patrimonio&"','"&cd_equipamento&"','"&cd_marca&"','"&cd_especialidade&"','"&cd_ns&"','1'"%>
						<%else%>
							<img src="../../imagens/px.gif" alt="" width="1" height="1" border="0" id="myImage" onload="alimenta_info_pat('','','','','','','');">
						<%end if%>						
					<!--/span-->
				</td>
			</tr>
			<%'	rs.close
			  '	set rs = nothing
			  if reg_conta = 0 then%>
				
			   <tr><td>O patrimômio/numero de série informado não existe.</td></tr>
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