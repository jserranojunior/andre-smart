<%sessao = strsession

if sessao = "" Then
sessao = "2" 
End if

sair = request("sair")
'cd_codigo  = 46
%>
<span style="border:1px solid silver; position:absolute; width:110px; height:97%; left:5px; top:20px; margin:0px; padding:0px; align: center; background:#d5d5d5; z-index:1;" class="no_print">

				<%''*** Usuário do sistema ***
				'*** Permissões ***
				xsql_usuario = "SELECT * FROM TBL_ADM_usuario_permissoes Where cd_usuario="&cd_codigo&""
				Set rs_permissao = dbconn.execute(xsql_usuario)
				
				menu1 = rs_permissao("menu_1")
				menu2 = rs_permissao("menu_2")
				menu3 = rs_permissao("menu_3")%>
	<ul id="nav" style=" margin-left:0px; margin-right:auto;">
		<li><a href="home.asp">Home</a></li> 
	    <%'*** Item principal do menu ***
				xsql_m1 ="SELECT * FROM tbl_menu_1 "
				Set rs_menu_1 = dbconn.execute(xsql_m1)
				
				while not rs_menu_1.EOF
				cd_menu_1 = rs_menu_1("cd_codigo")
				item_menu_1 = rs_menu_1("item_menu")
				link_menu_1 = rs_menu_1("link_menu")
				link_parametro_1 = rs_menu_1("link_parametro")
				
				mprincipal = instr(1,menu1,cd_menu_1,0)
					if mprincipal <> "0" then%>
					<li><a href="<%=link_menu_1%><%=link_parametro_1%>"><%=cd_menu_1%> - <%=item_menu_1%></a>
						<%'*** Item secundário do menu ***
								xsql_m2 ="SELECT * FROM tbl_menu_2 Where menu_principal = "&cd_menu_1&""
								Set rs_menu_2 = dbconn.execute(xsql_m2)%>
									<%if Not rs_menu_2.EOF Then%>
										<ul> 
									        <%do while not rs_menu_2.EOF
											cd_menu_2 = rs_menu_2("cd_codigo")
											item_menu_2 = rs_menu_2("item_menu")
											link_menu_2 = rs_menu_2("link_menu")
											link_parametro_2 = rs_menu_2("link_parametro")
											
											msecundario = instr(1,menu2,cd_menu_2,0)
											if msecundario <> 0 then%>
												<li style="border:1px solid #a7a7a7;"><a href="<%=link_menu_2%><%=link_parametro_2%>"><%'=cd_menu_1&"."&cd_menu_2&" - "%><%=item_menu_2%></a>
													<%'*** Item terciário do menu ***
															xsql_m3 ="SELECT * FROM tbl_menu_3 Where menu_principal = "&cd_menu_1&" AND menu_medio = "&cd_menu_2&""
															Set rs_menu_3 = dbconn.execute(xsql_m3)%>
															<%if Not rs_menu_3.EOF Then%>
																<ul>
																	<%while not rs_menu_3.EOF
																	cd_menu_3 = rs_menu_3("cd_codigo")
																	item_menu_3 = rs_menu_3("item_menu")
																	link_menu_3 = rs_menu_3("link_menu")
																	link_parametro_3 = rs_menu_3("link_parametro")
																	
																	mterciario = instr(1,menu3,cd_menu_3,0)
																		if mterciario <> 0 then%>
																			<li style="border:1px solid #676767;"><a href="<%=link_menu_3%><%=link_parametro_3%>"><%'=cd_menu_1&"."&cd_menu_2&"."&cd_menu_3&" - "%><%=item_menu_3%></a></li>
																		<%end if
																	rs_menu_3.movenext
																	wend%>																	
																</ul>																
															<%end if%>											
												</li>
											<%end if
											rs_menu_2.movenext
											loop%>
										</ul>
									<%end if%>
					</li>
					<%end if
				rs_menu_1.movenext
				wend%>		   
	 </ul>
	<br id="no_print">
<a href="default.asp?sair=1" id="no_print" style="width:10px; position:absolute; left:10px;">Sair</a>
<br id="no_print"><br id="no_print">
<!--img src="imagens/ic_print.gif" alt="imprimir" width="24" height="26" border="0" onclick="visualizarImpressao();" style="position:absolute; left:10px;" id="no_print"-->
&nbsp;<img src="imagens/ic_print.gif" alt="imprimir" width="24" height="26" border="0" onclick="visualizarImpressao();" id="no_print">
<!--#include file="calendario/calendario.asp"--></span>
<%'if int(session_cd_usuario) = int(46) then%>

<%'end if%>

