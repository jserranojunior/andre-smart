<%	'xsql = "Select * From View_funcionario_busca where datediff(m,dt_rgproval = 1 and dt_desliga is null"
	xsql = " SELECT     *, DATEDIFF(m,GETDATE(),dt_rgproval) AS validade FROM VIEW_funcionario_busca WHERE     (dt_desliga IS NULL) AND (DATEDIFF(m, GETDATE(),dt_rgproval) <= '1') order by validade"	
	'Set rs = dbconn.execute(xsql)"
	Set rs = dbconn.execute(xsql)%>


<table width="500" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" align=center >


	<%if not rs.eof then%>		
			<table width="98%" border="0" cellspacing="0" cellpadding="0" class="textoPadrao">
				<tr>
					<td class="txt_cinza" colspan="100%" id="no_print"><b>Lista &raquo; <font color="red">Colaboradores <%'=Ucase(str_func)&"S"%> &nbsp; <%=nm_regime_trab%></font></b> &nbsp;<br><br>/ <a href="empresa.asp?tipo=novo">Novo</a>/</td>
				</tr>
				<tr id="no_print">
					<td colspan="100%">&nbsp;&nbsp;<!--(Listagem: <a href="empresa.asp?tipo=lista&func=ativo">Ativos</a> / <a href="empresa.asp?tipo=lista&func=inativo">Inativos</a>)
					&nbsp;&nbsp;&nbsp; (<a href="empresa.asp?tipo=novo">Novo</a>)</td-->
				</tr>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr bgcolor=#c0c0c0>
					<td colspan="100%" align="center" style="color:white; font-size:12px;"><b>VENCIMENTO DOS CORENS DOS COLABORADORES</b></td>						
				 </tr>
				 <tr><td>&nbsp;</td></tr>
				 <%num = 1
				 
				 row_color = 1
				 
				 	
				 
				 
			Do While Not rs.eof 
			
				if row_color = 1 then
						cor_linha = "#FFFFFF"
					else
						cor_linha = "#ccffff"
					end if
				
				strcod = rs("cd_codigo")
				cd_regime_trabalho = rs("cd_regime_trabalho")
				cd_matricula = rs("cd_matricula")
				nm_nome = rs("nm_nome")
				nm_endereco = rs("nm_endereco")
				nr_numero = rs("nr_numero")
						
					
				nm_complemento = rs("nm_complemento")
				nm_bairro = rs("nm_bairro")
				nm_cep = rs("nm_cep")
				nm_cidade = rs("nm_cidade")
				nm_estado = rs("nm_estado")
				nm_endereco = nm_endereco&", "&nr_numero & nm_complemento &" <br> "&nm_bairro &" - "& nm_cep &" - "&nm_cidade &" - "& nm_estado
				
				'nm_sobrenome = rs("nm_sobrenome")
				contrato = rs("dt_contratado")
					if contrato <> "" Then
						contrato = zero(day(contrato))&"/"&zero(month(contrato))&"/"&right(year(contrato),2)
					end if
				desliga = rs("dt_desliga")
					if desliga <> "" Then
						desliga = zero(day(desliga))&"/"&zero(month(desliga))&"/"&right(year(desliga),2)
					end if
				
					xsql_cargo = "Select * From View_funcionario_cargo Where cd_funcionario = "&strcod&" ORDER BY dt_atualizacao desc"
					Set rs_cargo = dbconn.execute(xsql_cargo)
						if not rs_cargo.EOF then
							nm_cargo = rs_cargo("nm_cargo")
						end if
				
					xsql = "Select * From View_funcionario_horario Where cd_funcionario = "&strcod&" order by dt_atualizacao DESC"
					Set rs_hora = dbconn.execute(xsql)
						if not rs_hora.EOF Then
							hr_entrada = zero(hour(rs_hora("hr_entrada")))&":"&zero(minute(rs_hora("hr_entrada")))
							hr_saida = zero(hour(rs_hora("hr_saida")))&":"&zero(minute(rs_hora("hr_saida")))
							nm_intervalo = rs_hora("nm_intervalo")
						end if
						
					xsql = "Select * From View_funcionario_unidade Where cd_funcionario = "&strcod&" order by dt_atualizacao DESC"
					Set rs_unid = dbconn.execute(xsql)
						if not rs_unid.EOF Then
							nm_unidade = rs_unid("nm_unidade")
							nm_sigla = rs_unid("nm_sigla")
						end if
						
					xsql = "Select * From View_funcionario_status Where cd_funcionario = "&strcod&" order by dt_atualizacao DESC"
					Set rs_status = dbconn.execute(xsql)
						if not rs_status.EOF then
							nm_status = rs_status("nm_status")
						end if
					nm_regime_trab = rs("nm_regime_trab")
					cd_numreg = rs("cd_numreg")
					nm_rg = rs("nm_rg")
					nm_cpf = rs("nm_cpf")
					'nm_tit_eleitor = rs("nm_tit_eleitor")
					dt_rgproinscr = rs("dt_rgproinscr")
					dt_rgproval = rs("dt_rgproval")
					obs_rgpro = rs("obs_rgpro")
					validade = rs("validade")
					%>
				
				<%if validade <= 0 and cab_cor = "" then
				cab_cor = cab_cor + 1%>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr bgcolor="gray"><td colspan="100%" style="color:yellow; font-size:12px;" align="center"><b>CORENs vencidos</b></td></tr>
				<tr bgcolor=#c0c0c0>
					<td>&nbsp;</td>
					<td><b>Mat.</b></td>
					<td>&nbsp;<a href="empresa.asp?tipo=lista&ordem="><b>Nome</b></a></td>
					<!--td>&nbsp;<b>Endere�o</b></td>
					<td>&nbsp;<b>RG</b></td-->
					<td>&nbsp;<b>CPF</b></td>
					<td>&nbsp;<b>Coren</b></td>
					<td>&nbsp;<b>Inscri��o</b></td>
					<td>&nbsp;<b>Validade</b></td>
					<td>&nbsp;<b>Vencidos h�</b></td>
					<td>&nbsp;<b>Cargo</b></td>
					<td>&nbsp;<b>Empresa</b></td>
					<!--td>&nbsp;<b>Horario</b></td-->
					<td>&nbsp;<b>Unid.</b></td>
					<!--td>&nbsp;<b>Situa��o</b></td-->					
					<!--td>&nbsp;<b>Admiss�o</b class="textopadrao"></td-->
					<!--td>&nbsp;<b>Demiss�o</b></td-->
					<td>&nbsp;<b>Observa��es</b></td>
				 </tr>
				<%elseif validade > 0 and cab_cor = "1" then
				cab_cor = cab_cor + 1%>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr bgcolor="gray"><td colspan="100%" colspan="100%" style="color:white; font-size:12px;" align="center"><b>CORENs pr�ximos ao vencimento</b></td></tr>
				<tr bgcolor=#c0c0c0>
					<td>&nbsp;</td>
					<td><b>Mat.</b></td>
					<td>&nbsp;<a href="empresa.asp?tipo=lista&ordem="><b>Nome</b></a></td>
					<!--td>&nbsp;<b>Endere�o</b></td>
					<td>&nbsp;<b>RG</b></td-->
					<td>&nbsp;<b>CPF</b></td>
					<td>&nbsp;<b>Coren</b></td>
					<td>&nbsp;<b>Inscri��o</b></td>
					<td>&nbsp;<b>Validade</b></td>
					<td>&nbsp;<b>Vence em</b></td>
					<td>&nbsp;<b>Cargo</b></td>
					<td>&nbsp;<b>Empresa</b></td>
					<!--td>&nbsp;<b>Horario</b></td-->
					<td>&nbsp;<b>Unid.</b></td>
					<!--td>&nbsp;<b>Situa��o</b></td>					
					<td>&nbsp;<b>Admiss�o</b class="textopadrao"></td>
					<td>&nbsp;<b>Demiss�o</b></td-->
					<td>&nbsp;<b>Observa��es</b></td>
				 </tr>
				<%end if%>				
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs("cd_codigo")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>">
					<td style="color:silver;""><%=zero(num)%></td>
					<td valign="middle" align=right><%=cd_matricula%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td valign="middle">&nbsp;<%=nm_nome%></td>
					<!--td valign="middle"><%=nm_endereco%></td>
					<td valign="middle">&nbsp;<%=nm_rg%></td-->
					<td valign="middle">&nbsp;<%=nm_cpf%></td>
					
					<td valign="middle">&nbsp;<%=cd_numreg%></td>
					<td valign="middle">&nbsp;<%=dt_rgproinscr%></td>
					<td valign="middle">&nbsp;<%=dt_rgproval%></td>
					<%if validade < (-12) then
						if validade < 13 then
							tempo_ano = "ano"
						else
							tempo_ano = "anos"
						end if 
						'validade = mid(validade/12,1,instr(validade/12,","))&" "&tempo_ano
						'	validade = replace(validade,"-","")
							'validade = mid((validade/12),instr(validade,","),len(validade))
							validade_mes = mid((validade/12),instr(validade/12,",")+1,left(len(validade/12),1))
							
							validade = mid(validade/12,1,instr(validade/12,","))&validade_mes&" "&tempo_ano
							validade = replace(validade,"-","")
					else
						validade = replace(validade,"-","")
						if validade = 1 then
							tempo_mes = "mes"
						else
							tempo_mes = "meses"
						end if 
						
						validade = replace(validade,"-","")&" "&tempo_mes
						'validade = validade&"meses"
					end if
						'aviso_cor = %>
					
					
					<td valign="middle">&nbsp;<%=validade%></td>
					<!--td valign="middle">&nbsp;<%=nm_tit_eleitor%></td-->
					<td valign="middle">&nbsp;<%=nm_cargo%></td>
					<td valign="middle">&nbsp;<%=nm_regime_trab%></td>
					<!--td valign="middle">&nbsp;<%=hr_entrada%> - <%=hr_saida%></td-->
					<td valign="middle">&nbsp;<%=nm_sigla%><%'=nm_unidade%></td>
					<!--td valign="middle">&nbsp;<%=nm_status%></td>
					<td valign="middle" align="center"><%=contrato%></td>
					<td valign="middle" align="center"><%=desliga%></td-->
					<td valign="middle" align="center">&nbsp;<%=obs_rgpro%></td>					
				 </tr>				 
				 <%
				  nm_unidade = ""
				  nm_sigla = ""
				  nm_status = ""
				  cd_status = ""
				  strcor = ""
				  
				  nm_cargo = ""
				  hr_entrada = ""
				  hr_saida = ""
				  
				  
				  num = num + 1
				  
				  row_color = row_color + 1
				  	if row_color > 2 then
						row_color = 1
					end if
				  
					rs.movenext
					Loop
					rs.close
					Set rs = Nothing
					%>
				<tr><td>&nbsp;</td></tr>			
				<tr>
				 <td  colspan=5 width="98%" bgcolor=#FFFFFF><td>
				</tr>
				<tr>
				 	<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="250" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="140" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
					<!--td><img src="imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="10" height="1" border="0"></td-->
				</tr>
		</table> 
  <%End if%>
					</td>
				</tr>								
				<tr>
					<td height="26"><img src="../../imagens/px.gif" alt="" width="100" height="26" border="0"></td>
				</tr>
<%'rs_contrato.movenext
'wend%>				