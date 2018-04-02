<br>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="600">
			<tr bgcolor="#808080"><td colspan="100%" align="center"><b>MANUTENÇÃO DE LINKS E PARÂMETROS DO MENU</b></tr>
			<tr>
				<td colspan="100%">&nbsp;</td>				
			</tr>
			<%xsql_upermissoes ="SELECT * FROM tbl_ADM_usuario_permissoes where cd_usuario='"&cd_cod_usuario&"'"
			Set rs_upermissoes = dbconn.execute(xsql_upermissoes)
			if not rs_upermissoes.EOF Then
			permissoes_1 = rs_upermissoes("menu_1")
			permissoes_2 = rs_upermissoes("menu_2")
			permissoes_3 = rs_upermissoes("menu_3")
			End if%>
			
			<%'****************************
			'*** Item principal do menu ***
			'******************************
			
			xsql_m1 ="SELECT * FROM tbl_menu_1 "
			Set rs_menu_1 = dbconn.execute(xsql_m1)%>			
					<%do while not rs_menu_1.EOF
					cd_menu_1 = rs_menu_1("cd_codigo")
					item_menu_1 = rs_menu_1("item_menu")
					link_menu_1 = rs_menu_1("link_menu")
					link_parametro_1 = rs_menu_1("link_parametro")
						if not link_parametro_1 = null then
							link_parametro_1 = replace(link_parametro_1,"&","<amp>")
						end if
					tbl_destino = "1"%>					
					<tr><td valign="middle"><td><b>(<a href="#" onClick="window.open('sistema_adm/menu/menu_itens.asp?destino=1','links','width=330,height=195')" alt="Inserir">+</a>)</b></td>
						<%perm_1 = instr(1,permissoes_1,cd_menu_1,0)%><%if perm_1 <> 0 then%><%ck_perm1="checked"%><%else%><%ck_perm1=""%><%end if%>					
						<td valign="top" width="25%"><img src="imagens/blackdot.gif" width="250" height="1"><br><%=cd_menu_1%> - <a href="#" onClick="window.open('sistema_adm/menu/menu_links.asp?destino=<%=tbl_destino%>&cd_menu=<%=cd_menu_1%>&item_menu=<%=item_menu_1%>&link=<%=link_menu_1%>&parametro=<%if link_parametro_1 <> "" then%><%response.write(replace(link_parametro_1,"&amp;",".amp."))%><%end if%>','links','width=420,height=200')" alt="Alterar"><b><%=item_menu_1%></b></a><br><font color="#c0c0c0"><%=link_menu_1%><%=link_parametro_1%><br></font></td>
						<td width="75%">
			
			<%'*****************************
			'*** Item secundário do menu ***
			'*******************************

						xsql_m2 ="SELECT * FROM tbl_menu_2 Where menu_principal = "&cd_menu_1&""
						Set rs_menu_2 = dbconn.execute(xsql_m2)%>
						<%'if Not rs_menu_2.EOF Then%>
							<table border="1" width="400" cellpadding="0" cellspacing="0" bordercolordark="#000000" bordercolorlight="#000000">
							<%do while not rs_menu_2.EOF
							cd_menu_2 = rs_menu_2("cd_codigo")
							item_menu_2 = rs_menu_2("item_menu")
							link_menu_2 = rs_menu_2("link_menu")
							link_parametro_2 = rs_menu_2("link_parametro")
							tbl_destino = "2"%>
							<tr bgcolor="#c0c0c0">
								<td>&nbsp;</td>
								<%perm_2 = instr(1,permissoes_2,cd_menu_2,0)%><%if perm_2 <> 0 then%><%ck_perm2="checked"%><%else%><%ck_perm2=""%><%end if%>	
								<td valign="top" width="50%"><%=cd_menu_1%>.<%=cd_menu_2%> - <a href="#" onClick="window.open('sistema_adm/menu/menu_links.asp?destino=<%=tbl_destino%>&cd_menu=<%=cd_menu_2%>&item_menu=<%=item_menu_2%>&link=<%=link_menu_2%>&parametro=<%if link_parametro_2 <> "" then%><%response.write(replace(link_parametro_2,"&amp;",".amp."))%><%end if%>','links','width=420,height=200')" alt="Alterar"><b><%=item_menu_2%></b></a><br><font color="#ffffff"><%=link_menu_2%><%=link_parametro_2%><br></font><img src="imagens/px.gif" width="250" height="1">
								</td>
								<td><b>(<a href="#" onClick="window.open('sistema_adm/menu/menu_itens.asp?destino=3&item_menu1=<%=item_menu_1%>&cd_menu=<%=cd_menu_1%>&item_menu2=<%=item_menu_2%>&cd_menu2=<%=cd_menu_2%>','links','width=330,height=250')" alt="Inserir">+</a>)</b></td>
								<td width="50%">
								
			<%'****************************
			'*** Item terciário do menu ***
			'******************************

									xsql_m3 ="SELECT * FROM tbl_menu_3 Where menu_principal = "&cd_menu_1&" AND menu_medio = "&cd_menu_2&""
									Set rs_menu_3 = dbconn.execute(xsql_m3)%>
									<%do while not rs_menu_3.EOF
									cd_menu_3 = rs_menu_3("cd_codigo")
									item_menu_3 = rs_menu_3("item_menu")
									link_menu_3 = rs_menu_3("link_menu")
									link_parametro_3 = rs_menu_3("link_parametro")
									tbl_destino = "3"%>
									<%perm_3 = instr(1,permissoes_3,cd_menu_3,0)%><%if perm_3 <> 0 then%><%ck_perm3="checked"%><%else%><%ck_perm3=""%><%end if%>
									<%=cd_menu_1%>.<%=cd_menu_2%>.<%=cd_menu_3%> - <a href="#" onClick="window.open('sistema_adm/menu/menu_links.asp?destino=<%=tbl_destino%>&cd_menu=<%=cd_menu_3%>&item_menu=<%=item_menu_3%>&link=<%=link_menu_3%>&parametro=<%if link_parametro_3 <> "" then%><%response.write(replace(link_parametro_3,"&amp;",".amp."))%><%end if%>','links','width=420,height=200')" alt="Alterar"><b><%=item_menu_3%></b></a><br><font color="#ffffff"><%=link_menu_3%><%=link_parametro_3%></font><br>
									
									<%rs_menu_3.movenext
										loop%>
									&nbsp;<img src="imagens/px.gif" width="250" height="1">
								</td>
							</tr>
							<%
							rs_menu_2.movenext
							loop%>
							<tr bgcolor="#c0c0c0"><td><b>(<a href="#" onClick="window.open('sistema_adm/menu/menu_itens.asp?destino=2&item_menu1=<%=item_menu_1%>&cd_menu2=<%=cd_menu_2%>&item_menu2=<%=item_menu_2%>&cd_menu=<%=cd_menu_1%>','links','width=330,height=220')" alt="Inserir">+</a>)</b></td></tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
					
					<%
					link_menu_2 = ""
					link_parametro_2 = ""
					
					rs_menu_1.movenext
					loop%>
					<tr><td></td><td>&nbsp;</td></tr>
			</table>