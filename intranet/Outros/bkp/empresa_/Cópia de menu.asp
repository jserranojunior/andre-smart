
<%
sessao = strsession

if sessao = "" Then
sessao = "2" 
End if

sair = request("sair")

%>
<!--Informações apenas para os desenvolvedores
Desenvolvedores - 2
- Marcelo
- André

Administração - 1
- Paul
- Luiz

Supervisão - 3
- Sandra

Agenda - 4
- Marli

Básico - 5
- Paula
//-->
			
	 <ul id="nav">
		    <li><a href="home.asp">Home</a></li>
			<%
			'Vefificação de string (m = módulo c = cooperativa nº= número de referência)
			'mc1 = instr(1,modulos,"ca1",0)
			'mc2 = instr(1,modulos,"cb1",0)
			'mc3 = instr(1,modulos,"cc1",0)
			'mc4 = instr(1,modulos,"cd1",0)
			'mc5 = instr(1,modulos,"ce1",0)
			'if mc1 <> 0 OR mc2 <> 0 OR mc3 <> 0 OR mc4 <> 0 OR mc5 <> 0 Then
			
			'*** Usuário do sistema ***
			'*** Permissões ***
			xsql_usuario = "SELECT * FROM TBL_ADM_usuario_permissoes Where cd_usuario="&cd_codigo&""
			Set rs_permissao = dbconn.execute(xsql_usuario)
			
			menu1 = rs_permissao("menu_1")
			menu2 = rs_permissao("menu_2")
			menu3 = rs_permissao("menu_3")
			
			'*** Item principal do menu ***
			xsql_m1 ="SELECT * FROM tbl_menu_1 "
			Set rs_menu_1 = dbconn.execute(xsql_m1)
			
			do while not rs_menu_1.EOF
			cd_menu_1 = rs_menu_1("cd_codigo")
			item_menu_1 = rs_menu_1("item_menu")
			link_menu_1 = rs_menu_1("link_menu")
			
			mprincipal = instr(1,menu1,cd_menu_1,0)
			
			if mprincipal <> "0" then
			%>
			<li><a href="<%=link_menu_1%>"><%=cd_menu_1%> - <%=item_menu_1%></a></li>
				
				<%'*** Item secundário do menu ***
				xsql_m2 ="SELECT * FROM tbl_menu_2 Where menu_principal = "&cd_menu_1&""
				Set rs_menu_2 = dbconn.execute(xsql_m2)%>
				<%if Not rs_menu_2.EOF Then%>
				<ul>
					<%do while not rs_menu_2.EOF
					cd_menu_2 = rs_menu_2("cd_codigo")
					item_menu_2 = rs_menu_2("item_menu")
					link_menu_2 = rs_menu_2("link_menu")
					
					msecundario = instr(1,menu2,cd_menu_2,0)
					
					if msecundario <> 0 then%>					
					<li><a href="<%=link_menu_2%>"><%=cd_menu_1%>.<%=cd_menu_2%> - <%=item_menu_2%></a> 						 
				       	<%'*** Item terciário do menu ***
						xsql_m3 ="SELECT * FROM tbl_menu_3 Where menu_principal = "&cd_menu_1&" AND menu_medio = "&cd_menu_2&""
						Set rs_menu_3 = dbconn.execute(xsql_m3)%>
						<%if Not rs_menu_3.EOF Then%>
						<ul>
							<%do while not rs_menu_3.EOF
							cd_menu_3 = rs_menu_3("cd_codigo")
							item_menu_3 = rs_menu_3("item_menu")
							link_menu_3 = rs_menu_3("link_menu")
							
							mterciario = instr(1,menu3,cd_menu_3,0)
							
							if mterciario <> 0 then%>
							<li><a href="<%=link_menu_3%>"><%=cd_menu_1%>.<%=cd_menu_2%>.<%=cd_menu_3%> - <%=item_menu_3%></a></li>							
				    		<%end if
							rs_menu_3.movenext
							loop%>
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
			loop%>			
<br>

<br>
&nbsp;<a href="default.asp?sair=1">Sair</a>
<br>&nbsp;

<%'**********************************************************************************
'*					Exemplo de item de menu											*
'************************************************************************************
'*					<%if ma1 <> 0 Then% >
'					<li><a href="#">Item principal</a> 
'								<ul> 
'				        			<li><a href="#">Sub-item 1</a></li>
'									<li><a href="#">Sub-item 2</a></li> 
'										<ul>
	'										<li><a href="#">Item final</a></li>
'										</ul>
'				    			</ul>
'					</li>
'					<%end if% >
'***********************************************************************************
%>