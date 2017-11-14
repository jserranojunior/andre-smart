<!--#include file="../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%cd_patrimonio = request("no_patrimonio")
'pos_pat = instr(1,cd_patrimonio,":",1)
nm_tipo = mid(cd_patrimonio,1,1)
'no_patrimonio = mid(cd_patrimonio,pos_pat+1,len(cd_patrimonio))
'no_patrimonio = mid(cd_patrimonio,3,len(cd_patrimonio))
no_patrimonio = mid(cd_patrimonio,3,len(cd_patrimonio))

if len(cd_patrimonio) > 2 Then
	if IsNumeric(no_patrimonio) = true Then
		xsql = "SELECT count(no_patrimonio) as conta_pat FROM View_patrimonio_lista Where no_patrimonio = "&no_patrimonio&" and nm_tipo = '"&nm_tipo&"'"
		Set rs = dbconn.execute(xsql)
		conta_pat = rs("conta_pat")
		
		
		xsql = "SELECT * FROM View_patrimonio_lista Where no_patrimonio = "&no_patrimonio&" and nm_tipo = '"&nm_tipo&"'"
		Set rs = dbconn.execute(xsql)%>
		
		<table width="380" align="left">
			<%while not rs.EOF
			'if not rs.EOF Then%>
				<%cd_pat = rs("cd_codigo")
				no_patrimonio = rs("no_patrimonio")
				cd_equipamento = rs("cd_equipamento")
				nm_equipamento = rs("nm_equipamento")
				cd_marca = rs("cd_marca")
				nm_marca = rs("nm_marca")
				cd_especialidade = rs("cd_especialidade")
				nm_especialidade = rs("nm_especialidade")
				cd_ns = rs("cd_ns")%>
									
				<%if reg_conta >= 1 then%>
				<tr><td colspan="5" style="background-color:gray;"></td></tr>
				<%end if%>
				<tr>
					<td rowspan="3" align="center" valign="middle">
						<%if int(conta_pat) > 1 then%><input type="radio" name="seleciona_patrimonio" value="" onMouseup="sel_patrimonio('<%=cd_pat%>');"><%end if%>
					</td>
					<td style="width:45px;"><b>Pat:</b></td>
					<td style="color:#339900;"><b><%=no_patrimonio%></b></td>					
				</tr>
			  	<tr>
					<td style="width:45px;"><b>Item:</b></td>
					<td><%=nm_equipamento%></td>
					<td style="width:45px;"><b>Especialidade:</b></td>
					<td><%=nm_especialidade%></td>			
				</tr>
				<tr>
					<td style="width:45px;"><b>Marca:</b></td>
					<td><%=nm_marca%></td>
					<td style="width:45px;"><b>N/S:</b></td>
					<td><%=cd_ns%></td>
				</tr>
			<%
			reg_conta = reg_conta + 1
			rs.MoveNext
			wend
			'end if%>
			<tr>
				<td colspan="4">
					<span id="divPat_reg">
						<input type="hidden" name="cd_patrimonio" value="<%if int(conta_pat) <= 1 then response.write(cd_pat)%>">
						<input type="hidden" name="no_patrimonio" value="<%if int(conta_pat) <= 1 then response.write(no_patrimonio)%>">
						<input type="hidden" name="cd_equipamento" value="<%if int(conta_pat) <= 1 then response.write(cd_equipamento)%>">
						<input type="hidden" name="cd_marca" value="<%if int(conta_pat) <= 1 then response.write(cd_marca)%>">
						<input type="hidden" name="cd_especialidade" value="<%if int(conta_pat) <= 1 then response.write(cd_especialidade)%>">
						<input type="hidden" name="cd_ns" value="<%if int(conta_pat) <= 1 then response.write(cd_ns)%>">
						<input type="hidden" name="num_qtd" value="1">
					</span>
				</td>
			</tr>
			<%rs.close
			  set rs = nothing
			  if reg_conta = 0 then%>
				
			   <tr><td>O patrimômio informado não existe.</td></tr>
			  <%end if%>
		</table>
	<%else%>
		<span style="color:red;">Digite apenas números.<%=cd_patrimonio%></span>
	<%end if%>	
<%Else%>
	Nenhum registro encontrado.
<%end if%>