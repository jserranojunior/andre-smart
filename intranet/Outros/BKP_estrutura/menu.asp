
<%
sessao = strsession

if sessao = "" Then
sessao = "2" 
End if

sair = request("sair")

%>
<!--Informa��es apenas para os desenvolvedores
Desenvolvedores - 2
- Marcelo
- Andr�

Administra��o - 1
- Paul
- Luiz

Supervis�o - 3
- Sandra

Agenda - 4
- Marli

B�sico - 5
- Paula
//-->
			
	 <ul id="nav">
		    <li><a href="home.asp">Home</a></li>
			<%
			'Vefifica��o de string (m = m�dulo c = cooperativa n�= n�mero de refer�ncia)
			mc1 = instr(1,modulos,"ca1",0)
			mc2 = instr(1,modulos,"cb1",0)
			mc3 = instr(1,modulos,"cc1",0)
			mc4 = instr(1,modulos,"cd1",0)
			mc5 = instr(1,modulos,"ce1",0)
			if mc1 <> 0 OR mc2 <> 0 OR mc3 <> 0 OR mc4 <> 0 OR mc5 <> 0 Then
			%>
			<li><a href="home.asp">Cooperativa</a></li>
					<ul>
					<%if mc1 <> 0 Then%>
					<li><a href="produtividade.asp">Produtividade</a></li><%end if%>
					<%if mc2 <> 0 Then%>
					<li><a href="beneficios.asp">Benef�cios</a></li><%end if%>
					<%if mc3 <> 0 Then%>
					<li><a href="#">Relat�rios</a> 
								<ul> 
				        			<li><a href="relatorio_coop.asp">Cooperativa Final</a></li>
									<li><a href="relatorio_coop.asp?relatorio=consolidado">Consolidado (Tela)</a></li> 
				    			</ul></li><%end if%>
					<%if mc4 <> 0 Then%>			
					<li><a href="#">Colaboradores</a></li>
								<ul> 
									<li><a href="coop_cooperados_lista.asp">Listagem</a></li>
									<li><a href="coop_cooperados_cad.asp">Novo</a></li>
								</ul>
					</li>
					<%end if%>
					<%if mc5 <> 0 Then%>
					<li><a href="#">Fechamento</a></li>
								<ul> 
									<li><a href="coop_fecha_mes.asp">Per�odo: M�s</a></li>
								</ul>
					</li><%end if%>			 
					</ul> 
			</li>
			<%end if%>
			<%
			'Vefifica��o de string (m = m�dulo e = empresa n�= n�mero de refer�ncia)
			me1 = instr(1,modulos,"ea1",0)
			me2 = instr(1,modulos,"eb1",0)
			me3 = instr(1,modulos,"ec1",0)
			me4 = instr(1,modulos,"ed1",0)
			me5 = instr(1,modulos,"ee1",0)
			me6 = instr(1,modulos,"ef1",0)
			if me1 <> 0 OR me2 <> 0 OR me3 <> 0 OR me4 <> 0 OR me5 <> 0 OR me6 <> 0 Then
			%>
			<li><a href="home.asp">CMI</a></li>
					<ul>
					<%if me4 <> 0 Then%>			
					<li><a href="empresa_funcionarios_lista.asp">Funcion�rios</a></li>
								<ul> 
									<li><a href="empresa_funcionarios_lista.asp">Listagem</a></li>
									<li><a href="empresa_funcionarios_cad.asp">Novo</a></li>
								</ul>
					</li>
					<%end if%>
					<%if me1 <> 0 Then%>
					<li><a href="empresa_funcionarios_horas.asp">Horas Trabalhadas</a></li><%end if%>
								<ul>
									<li><a href="empresa_funcionarios_horas.asp">Digita��o</a></li>
									<li><a href="empresa_funcionarios_abono.asp">Abono</a></li>
									<li><a href="empresa_funcionarios_obs.asp">Observa��es</a></li>
								</ul>
					<%if me6 <> 0 Then%>
					<li><a href="#">Relat�rios</a> 
								<ul> 
				        			<!--li><a href="banco_horas_cred.asp">Cr�dito (hora Extra)</a></li>
									<li><a href="#">D�bito (folga)</a></li-->
									<li><a href="empresa_banco_horas.asp">Banco de horas</a></li>
									<li><a href="empresa_funcionarios_faltas.asp">Faltas</a></li> 
				    			</ul></li><%end if%>
					<%if me2 <> 0 Then%>
					<li><a href="#">Benef�cios</a></li><%end if%>
					<%if me3 <> 0 Then%>
					<li><a href="#">Relat�rios</a> 
								<ul> 
				        			<li><a href="#">M�s Final</a></li>
									<li><a href="#">Consolidado (Tela)</a></li> 
				    			</ul></li><%end if%>
					<%'if me5 <> 0 Then%>
					<!--li><a href="#">Fechamento</a></li>
								<ul> 
									<li><a href="emp_fecha_mes.asp">Per�odo: M�s</a></li>
								</ul>
					</li--><%'end if%>		
					<%'if me10 <> 0 Then%>
					<li><a href="#">Administrativo</a></li>
								<ul> 
									<li><a href="adm_cargos.asp">Cargos</a></li>
									<li><a href="adm_ajuda_custo.asp">Ajuda de custo (2)</a></li>
									<li><a href="adm_insalubridade.asp">Insalubridade (3)</a></li>
									<li><a href="adm_ad_noturno.asp">Adicional noturno (4)</a></li>
									<li><a href="adm_aux_creche.asp">Aux�lio Creche (5)</a></li>
									<li><a href="adm_v_alim.asp">Vale alimenta��o (7)</a></li>
									<li><a href="adm_v_transp.asp">Vale transporte (8)</a></li>
									<li><a href="emp_fecha_mes.asp">Unidades</a></li>
									<li><a href="emp_fecha_mes.asp">Benef�cios</a></li>
									<li><a href="emp_fecha_mes.asp">Referenciais</a></li>
									<li><a href="emp_fecha_mes.asp">Fechamento</a></li>
								</ul>
					</li><%'end if%>	 
					</ul> 
			</li>
			<%end if%>
			<%
			'Vefifica��o de string (m = m�dulo m = manuten��o n�= n�mero de refer�ncia)
			mm1 = instr(1,modulos,"ma1",0)
			mm2 = instr(1,modulos,"mb1",0)
			mm3 = instr(1,modulos,"mc1",0)
			if mm1 <> 0 OR mm2 <> 0 OR mm3 <> 0 Then
			%>		
		     <li><a href="#">Manuten��o</a> 
					<ul><%if mm1 <> 0 Then%> 
				      	<li><a href="manutencao.asp">Manuten��o</a></li><%end if%>
						<%if mm2 <> 0 Then%> 
				        <li><a href="#">Cadastramento</a></li> 
								<ul>
									<li><a href="fornecedores.asp">Fornecedores</a></li>
									<li><a href="equipamentos.asp">Equipamentos</a></li>
									<!--li><a href="#">Especialidades</a></li-->
									<li><a href="marcas.asp">Marcas</a></li>
									<li><a href="unidades.asp">Unidades</a></li>
									
								</ul><%end if%>
						<%if mm3 <> 0 Then%> 
				        <li><a href="relatorio_manut.asp">Relat�rios</a></li><%end if%>
				    </ul> 
		    </li>
			<%end if%>
			<%
			'Vefifica��o de string (m = m�dulo p = protocolo n�= n�mero de refer�ncia)
			mp1 = instr(1,modulos,"pa1",0)
			mp2 = instr(1,modulos,"pb1",0)
			mp3 = instr(1,modulos,"pc1",0)
			if mp1 <> 0 OR mp2 <> 0 OR mp3 <> 0 Then
			%>	
			<li><a href="#">Protocolos</a> 
					<ul><%if mp1 <> 0 Then%>	 
				        <li><a href="protocolo.asp?tipo=ins">Digita��o</a></li>
						<li><a href="protocolo.asp?tipo=ver">Ver/Alterar</a></li><%end if%> 
						<%if mp2 <> 0 Then%>
				    	<li><a href="#">Relat�rios</a></li> 
								<ul>
									<li><a href="#">Emiss�o de Etiquetas</a></li>
									<li><a href="#">Equipamentos</a></li>
									<li><a href="#">Especialidades</a></li>
									<li><a href="#">Marcas</a></li>
								</ul>
						<%end if%>		
						<%if mp3 <> 0 Then%>
						<li><a href="#">Cadastros</a></li><%end if%>
				   </ul> 
		    </li>
			<%end if%>
			<%'Vefifica��o de string (m = m�dulo f = folha de pagamento n�= n�mero de refer�ncia)
			mf1 = instr(1,modulos,"fa1",0)
			if mf1 <> 0 Then
			%>
			<li><a href="#">Folha de Pagamento</a> 
					<ul><%if mf1 <> 0 Then%>	 
				        <li><a href="folha_pagamento.asp?tipo=ins">Assistente</a></li><%end if%> 
					</ul> 
		    </li>
			<%end if%>
			<%'Vefifica��o de string (m = m�dulo i = invent�rio n�= n�mero de refer�ncia)
			mi1 = instr(1,modulos,"ia1",0)
			if mi1 <> 0 Then
			%>
			<li><a href="javascript:;">Invent�rio</a> 
					<ul> 
				    	<li><a href="patrimonio_cadastro.asp">Cadastro de produtos</a></li>
						<li><a href="patrimonio_descarte.asp">Descarte de produtos</a></li>
					</ul>
			</li>
			<%end if%>
			<!--li><a href="javascript:;">Item principal</a> 
					<ul> 
				    	<li><a href="<%=sub_item%>">Sub-item 1</a></li>
							<ul>
								<li><a href="#">Item final</a></li>
							</ul>
						<li><a href="#">Sub-item 2</a></li> 
					</ul>
			</li-->
			
			<%
			'Vefifica��o de string
			ma1 = instr(1,modulos,"aa1",0)
			if ma1 <> 0 Then
			%>	
			<li><a href="#">ADM</a> 
					<ul><%if ma1 <> 0 Then%> 
				        <li><a href="adm.asp?cd_codigo=">Usuarios</a></li> <%end if%>
				    </ul> 
		    </li>
			<%End if%> 
			
</ul>
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