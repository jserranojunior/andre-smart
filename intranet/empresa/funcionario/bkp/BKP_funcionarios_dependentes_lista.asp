<%	xsql = "Select * From View_funcionario_dependentes order by dt_nascimento desc"
	Set rs = dbconn.execute(xsql)%>


<table width="500" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" align=center >


	<%if not rs.eof then%>		
			<table width="98%" border="0" cellspacing="0" cellpadding="0" class="textoPadrao">
				<!--tr>
					<td class="txt_cinza" colspan="100%"><b>Lista &raquo; <font color="red">Funcionarios <%'=Ucase(str_func)&"S"%> &nbsp; <%=nm_regime_trab%></font></b> &nbsp;<br><br>/ <a href="empresa.asp?tipo=novo">Novo</a>/</td>
				</tr-->
				<tr><td>&nbsp;</td></tr>
				<tr><td align="center" colspan="100%" style="font-size:15; color:red;" bgcolor="#e8e8e8">&nbsp;Listagem de Dependentes</td></tr>
				<tr id="no_print">
					<td colspan="100%">&nbsp;&nbsp;<!--(Listagem: <a href="empresa.asp?tipo=lista&func=ativo">Ativos</a> / <a href="empresa.asp?tipo=lista&func=inativo">Inativos</a>)
					&nbsp;&nbsp;&nbsp; (<a href="empresa.asp?tipo=novo">Novo</a>)</td-->
				</tr>
				<tr bgcolor=#c0c0c0>
					<td>&nbsp;</td>					
					<td>&nbsp;&nbsp;<b>Colaborador</b></td>
					<td>&nbsp;<b>Dependente</b></td>
					<td>&nbsp;<b>Data de Nascimento</b></td>
					<td>&nbsp;<b>Idade</b></td>
					<td>&nbsp;<b>Parentesco</b></td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>			
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
				strnm_funcionario = rs("nm_funcionario")
				strnm_nome = rs("nm_nome")
				strdt_nascimento = rs("dt_nascimento")
				strnm_parentesco = rs("nm_parentesco")
				strnm_obs = rs("nm_obs")
				str_idade = rs("idade")
				idade = datediff("m",strdt_nascimento,now)
						if idade >= 12 then
							if idade = 12 then
								tempo = "ano"
							else
								tempo = "anos"
							end if
							'idade = int(idade/12)&" "&tempo
							
							idade_mes = mid((idade/12),instr(idade/12,",")+1,left(len(idade/12),1))
							
							idade = mid(idade/12,1,instr(idade/12,","))&idade_mes&" "&tempo
							idade = replace(idade,"-","")
							
						else 
							if idade <= 1 then
								tempo = "mes"
							else
								tempo = "meses"
							end if
							idade = idade&" "&tempo
						end if%>
						
				<%if str_idade <= 83 and idade_cab = "" then%>
				<tr bgcolor="#008080"><td colspan="100%" style="color:white;" align="center">Dependentes beneficiados pelo auxílio creche (até 6 anos e 11 meses)</td></tr>
				<%elseif str_idade > 83 AND idade_cab = "1" then%>
					<%if xpto = "" Then%>
				<tr bgcolor="#008080"><td colspan="100%" style="color:white;" align="center">&nbsp;Dependentes apartir de 7 anos</td></tr>
					<%xpto = 1
					num=1
					end if%>
				<%end if%>
				
				 <tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa.asp?tipo=cadastro&pag=<%=strpag%>&cod=<%=rs("cd_codigo")%>&busca=<%=strbusca%>');" onmouseout="mOut(this,'<%=cor_linha%>');" style="height:20px;" bgcolor="<%=cor_linha%>">
					<td style="color:silver;""><%=zero(num)%></td>
					<td valign="middle">&nbsp;
					<%if strnm_funcionario <> func then%>
						<%=strnm_funcionario%>
					<%end if%></td>
					<td valign="middle"><%=strnm_nome%></td>
					<td valign="middle">&nbsp;<%=strdt_nascimento%></td>
					<td valign="middle">&nbsp;<%=idade%></td>
					<td valign="middle" align="center">&nbsp;<%=strnm_parentesco%></td>
					<td valign="middle">&nbsp;<%'=str_idade%></td>
					<td valign="middle">&nbsp;<%=strnm_obs%></td>									
				 </tr>				 
				 <%
				  'nm_unidade = ""
				  'nm_sigla = ""
				  'nm_status = ""
				  'cd_status = ""
				  'strcor = ""
				 ' 
				 ' nm_cargo = ""
				 ' hr_entrada = ""
				 ' hr_saida = ""
				 func = strnm_funcionario
				  
				  num = num + 1
				 if str_idade < 83 then
				 	idade_cab = idade_cab + 1
				elseif str_idade >= 83 then
					idade_cab = 1
				 end if
				 
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
				 <td  colspan=5 width="98%" bgcolor=#c0c0c0><td>
				</tr>
				<tr>
				 	<td><img src="imagens/px.gif" alt="" width="20" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="250" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					<td><img src="imagens/px.gif" alt="" width="100" height="1" border="0"></td>
					
					
				</tr>
		</table> 
  <%End if%>
					</td>
				</tr>								
				<tr>
					<td height="26"><img src="px.gif" alt="" width="1" height="26" border="0"></td>
<%'rs_contrato.movenext
'wend%>				