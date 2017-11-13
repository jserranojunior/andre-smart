		<%if cd_cod_protocolo = "" then
			acao = "inserir"
		else
			acao = "editar"
		end if%>
		
		<%'if cd_user = 46 then
			mostra_teste = 0
		'end if%>
		
		
		<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
		<input type="hidden" name="no_protocolo"  value="<%=no_protocolo%>">
		<input type="hidden" name="no_digito" value="<%=no_digito%>">
		<input type="hidden" name="cd_cod_protocolo" value="<%=cd_cod_protocolo%>">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cd_libera_data" value="<%=cd_libera_data%>">
		
		<tr bgcolor="#f5f5f5">
			<td colspan="4">&nbsp;Protocolo: <b><%=digitov(zero(int(cd_unidade))&"."&proto(no_protocolo))%> &nbsp;<%=mensagem%></b><%'=cd_cod_protocolo%></td>
		</tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="4"><font color="white"><b>&nbsp;Identificação do Paciente - Etiqueta</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td align="left" colspan="3">&nbsp;Nome</td>
			<td align="left" colspan="1">&nbsp;Idade</td>
		</tr>
		<!--tr>
			<td colspan="10"-->
			<!--/td>
		</tr-->			
		<tr bgcolor="#f7f7f7">
			<td align="left" colspan="3"><b><input type="text" name="nm_nome" id="nm_nome" size="60" maxlength="60" class="inputs" value="<%=nm_nome%>" onkeypress="detecta_enter(event,'cd_idade','<%=cd_user%>')" id="demo"></b>
			<img src="../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="foca_nome();"></td>
			<td align="left"><b><input type="text" name="cd_idade" id="cd_idade" size="5" maxlength="50" class="inputs" value="<%=cd_idade%>" onkeypress="detecta_enter(event,'nm_registro','<%=cd_user%>')"></b></td>
		</tr>
		<tr bgcolor="#f7f7f7"><td colspan="4">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td align="left" colspan="1">&nbsp;Registro</td>
			<td align="left" colspan="3">&nbsp;Convênio: </td>
			<!--td>&nbsp;Cirurgia Realizada?</td-->
		</tr>
		<tr bgcolor="#f7f7f7">
			<td align="left" ><b><input type="text" name="nm_registro" id="nm_registro" size="20" maxlength="20" class="inputs" value="<%=cd_registro%>" onkeypress="detecta_enter(event,'cd_convenio','<%=cd_user%>')"></b>&nbsp;&nbsp;
			<td align="left" valign="top" colspan="3">
			<b><%'strsql ="Select * from TBL_convenio order by nm_convenio"
						  		'Set rs_conv = dbconn.execute(strsql)%>
				<input type="text" name="cd_convenio" id="cd_convenio" size="3" maxlength="15" class="inputs" value="<%=cd_convenio%>"  onkeypress="detecta_enter(event,'nm_agenda_cirurgia','<%=cd_user%>')" onKeyup="mostra_convenio();">
				<input type="hidden<%if mostra_teste = 1 then response.write("s")%>" name="cd_convenio_x" id="cd_convenio_x" size="3" maxlength="15" value="<%=cd_convenio%>">
			<span id='divAssist_m1'>
				<select name="cd_convenio_1" id="cd_convenio_1" class="inputs" onChange="alimenta_convenio(this.value);" onBlur="alimenta_convenio(this.value);" onkeypress="detecta_enter(event,'nm_agenda_cirurgia','<%=cd_user%>')">
					<option value=""></option>
					<%strsql ="Select * from TBL_convenios where cd_status = 1 order by nm_convenio"
					Set rs_conv = dbconn.execute(strsql)
						Do while Not rs_conv.EOF
						'if Not rs_conv.EOF Then%>
						<%cd_conv = rs_conv("cd_codigo")%>
						<%if cd_convenio=cd_conv Then%><%ck_conv="selected"%><%else ck_conv=""%><%end if%>
						<option value="<%=rs_conv("cd_codigo")%>" <%=ck_conv%>><%=rs_conv("nm_convenio")%></option><%rs_conv.movenext
						Loop
						'end if
						rs_conv.close
						set rs_conv = nothing%>
				</select>
			</span>
			<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/convenios_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=160'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->	
			</td>
			<input type="hidden" name="nm_cirug_realizada" value="S">
			<!--td> <!-- *** Opção de selecionar a confirmação da cirurgia foi retirada a pedido da Marli em 27/10/2015, André***
				<select name="nm_cirug_realizada" id="nm_cirug_realizada" class="inputs" onkeypress="detecta_enter(event,'dt_dia_cirurgia','<%=cd_user%>')">
					<option value=""></option>
					<option value="S" <%=ck_real_s%><%if nm_cirug_realizada="" OR nm_cirug_realizada="S" Then response.write("selected")%>>SIM</option>
					<option value="N" <%=ck_real_n%><%if nm_cirug_realizada="N" Then response.write("selected")%>>NÃO</option>
				</select>
			</td-->				
		</tr>
		<tr bgcolor="#f7f7f7"><td colspan="4">&nbsp;<!--Pula linha 1--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="4"><font color="white"><b>&nbsp;Agenda</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="2">&nbsp;Agendamento</td>
			<td align="left">&nbsp;Sexo: </td>
			<td>&nbsp;Realizada no C.O.</td>
		</tr>
		<tr>
			<td colspan="2"><b>
				<select name="nm_agenda_cirurgia" id="nm_agenda_cirurgia" class="inputs" onFocus="nextfield ='dt_hora_agenda';" onkeypress="detecta_enter(event,'dt_hora_agenda','<%=cd_user%>')">
					<option value=""></option>
					<%if nm_agenda="A" Then%><%ck_agd_a = "selected"%>
					<%Elseif nm_agenda="E" Then%><%ck_agd_e = "selected"%>
					<%Elseif nm_agenda="N" Then%><%ck_agd_n = "selected"%>
					<%Elseif nm_agenda="P" Then%><%ck_agd_p = "selected"%>
					<%Elseif nm_agenda="U" Then%><%ck_agd_u = "selected"%>
					<%end if%>
					<option value="N" <%=ck_agd_n%>>Normal</option>
					<option value="A" <%=ck_agd_a%>>A seguir</option>
					<option value="E" <%=ck_agd_e%>>Empréstimo</option>					
					<option value="P" <%=ck_agd_p%>>Pós Agendada</option>
					<option value="U" <%=ck_agd_u%>>Urgência</option>					
				</select>
				</b>
			</td>
			<td align="left"><b><!--select name="nm_sexo" size="1" class="inputs" onFocus="nextfield ='nm_agenda_cirurgia';"-->
							<select name="nm_sexo" id="nm_sexo" size="1" class="inputs" onkeypress="detecta_enter(event,'nm_cirug_realizada','<%=cd_user%>')">
							 <option value=""></option>
							 <%if nm_sexo="M" Then%><%ck_m = "selected"%>
							 <%Elseif nm_sexo="F" Then%><%ck_f = "selected"%>
							 <%end if%>
							 <option value="M" <%=ck_m%>>1 Masculino</option>
							 <option value="F" <%=ck_f%>>2 Feminino</option>&nbsp;&nbsp;
							 </select>
							 </b></td>
			<td>&nbsp;<input type="checkbox" name="cd_co" id="cd_co" value="1" <%=co_ck%> onMouseup="detecta_enter(event,'cd_crm','<%=cd_user%>')">&nbsp;Sim</td>
		</tr>
		<tr bgcolor="#f7f7f7"><td colspan="4">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">			
			<td colspan="1">&nbsp;Data</td>
			<td colspan="1">&nbsp;Hora agendada</td>
			<td>&nbsp;Inicio</td>
			<td>&nbsp;Término</td>
		</tr>
		<tr  bgcolor="#f7f7f7">
			<td colspan="1">
				<%if dt_procedimento <> "" Then
						dt_proced_dia = DAY(dt_procedimento)
						dt_proced_mes = MONTH(dt_procedimento)
						dt_proced_ano = YEAR(dt_procedimento)
					else
						'dt_proced_dia = DAY(now)
						'dt_proced_mes = MONTH(now)
						dt_proced_ano = YEAR(now)
						
					end if%><b>
							<input type="text" name="dt_dia_cirurgia" id="dt_dia_cirurgia" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(dt_proced_dia)%>">
							<input type="text" name="dt_mes_cirurgia" id="dt_mes_cirurgia" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(dt_proced_mes)%>" onBlur="verifica_data();"  onKeyup="verifica_data();" onkeypress="detecta_enter(event,'dt_ano_cirurgia','<%=cd_user%>')">
							<input type="text" name="dt_ano_cirurgia" id="dt_ano_cirurgia" size="4" maxlength="4" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=dt_proced_ano%>" onkeypress="detecta_enter(event,'dt_hora_agenda','<%=cd_user%>')"></b>
			</td>
			<td><%if dt_hr_agenda <> "" Then
					dt_min_agenda = minute(dt_hr_agenda)
					dt_hr_agenda = hour(dt_hr_agenda)
				end if
				if dt_inicio <> "" Then
					dt_min_inicio = minute(dt_inicio)
					dt_hr_inicio = hour(dt_inicio)
				end if
				if dt_fim <> "" Then
					dt_min_fim = minute(dt_fim)
					dt_hr_fim = hour(dt_fim)
				end if%>
				<b><input type="text" size="2" name="dt_hora_agenda" id="dt_hora_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_hr_agenda)%>"> :				
				<input type="text" size="2" name="dt_minuto_agenda" id="dt_minuto_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_min_agenda)%>"></b>
			</td>
			<td colspan="1">&nbsp;<b><input type="text" size="2" name="dt_hora_inicio" id="dt_hora_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_hr_inicio)%>"> :
				<input type="text" size="2" name="dt_minuto_inicio" id="dt_minuto_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_min_inicio)%>"></b></td>
			<td>&nbsp;<b><input type="text" size="2" name="dt_hora_fim" id="dt_hora_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_hr_fim)%>"> :
				<input type="text" size="2" name="dt_minuto_fim" id="dt_minuto_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_min_fim)%>"></b></td>
			
			
		</tr>
		<tr bgcolor="#f7f7f7"><td colspan="4">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="4" style="color:white;"><b>&nbsp;Procedimento Cirúrgico</b></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="4">&nbsp;CRM &nbsp; / Médico</td>
		</tr>

		<tr bgcolor="#f7f7f7">
			<%strsql ="Select * from TBL_Medicos Where A055_codmed='"&cd_med&"' and a055_status = 1"
					Set rs_med = dbconn.execute(strsql)
					do while not rs_med.EOF 
						cd_crm = rs_med("A055_crmmed")
					rs_med.movenext
					loop
					
					rs_med.close
					set rs_med = nothing
					'wend' if%>
			<td colspan="4">&nbsp;<b><input type="text" name="cd_crm" id="cd_crm" class="inputs" size="10" maxlength="30" value="<%=cd_crm%>" onKeyup="mostra_medico();" onkeypress="detecta_enter(event,'cd_rack','<%=cd_user%>')">
				<input type="hidden" name="cd_crm_x" id="cd_crm_x" size="5" maxlength="30" value="<%=cd_med%>">
		<span id='divMed'><select name="cd_crm_1" class="inputs" onChange="alimenta_medico(this.value);" onBlur="alimenta_medico(this.value);" onFocus="nextfield ='cd_rack';" onkeypress="detecta_enter(event,'cd_rack','<%=cd_user%>')">
								<option value=""></option>
								<%'strsql ="Select * from TBL_medicos where A055_codmed like '%"&cd_med&"' AND a055_status = 1 order by A055_desmed"
								strsql ="Select * from TBL_medicos where a055_status = 1 order by A055_desmed"
								Set rs_med = dbconn.execute(strsql)
								if not rs_med.EOF Then
								'cd_crm = rs_med("A055_crmmed")
								end if%>
								<%Do while Not rs_med.EOF%>
								<%cd_medico = rs_med("A055_codmed")
								cd_crm = rs_med("A055_crmmed")
								
									'if cd_user = 46 then
										check_medico = instr(1,cd_med,cd_medico,1)
									'end if%>
								<%'if cd_medico = cd_med Then%>
								<%'ck_med="selected"%><%'else ck_med=""%>
								<%'end if%>
								<option value="<%=cd_medico%>" <%if check_medico = 1 then response.write("selected")%>><%'=cd_medico&"/"&cd_med%><%=rs_med("A055_desmed")%> <%'=check_medico%></option><%rs_med.movenext
						 		Loop
								
								rs_med.close
								set rs_med = nothing%>
							</select>
			</span></b>
			<a href="#" onclick="window.open('janelas_cadastros/medicos_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=160'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a>
			</td>							
		</tr>


		<tr bgcolor="#f7f7f7"><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="2">&nbsp;Rack</td>
			<td colspan="2">&nbsp;Especialidade</td>
			<!--td>&nbsp;</td-->
		</tr>
		<tr bgcolor="#f7f7f7">
			<td colspan="2">&nbsp;<b><input type="text" name="cd_rack" id="cd_rack" size="3" class="inputs" value="<%=cd_rack%>" onKeyup="mostra_rack();" onkeypress="detecta_enter(event,'cd_especialidade','<%=cd_user%>')">
									<input type="hidden" name="cd_rack_hide" value="">
									<input type="hidden<%if mostra_teste = 1 then response.write("s")%>" name="cd_rack_x" id="cd_rack_x" value="<%=cd_rack%>" size="3">
			<span id='divRack'><select name="cd_rack_1" id="cd_rack_1" class="inputs" onChange="alimenta_rack(this.value);" onBlur="alimenta_rack(this.value);" onkeypress="detecta_enter(event,'cd_especialidade','<%=cd_user%>')">
								<option value=""></option>
							<%if cd_unidade = 20 Then
								strsql ="Select * from TBL_Rack where a056_status = 1 and a056_codrac between 11 and 14 order by A056_desrac"' Força mostrar apenas os racks do Hosp. Edmundo
							Else
								strsql ="Select * from TBL_Rack where a056_status = 1 order by A056_desrac"' order by nm_rack"
							end if
						  		Set rs_rack = dbconn.execute(strsql)%>
								<%Do while Not rs_rack.EOF%><%cd_rk = rs_rack("A056_codrac")%>
								<%if cd_rack=int(cd_rk) Then%><%ck_rack="selected"%><%else ck_rack=""%><%end if%>
								<option value="<%=rs_rack("A056_codrac")%>" <%=ck_rack%>><%=rs_rack("A056_desrac")%></option><%rs_rack.movenext
						 		Loop
								
								rs_rack.close
								set rs_rack = nothing%>
							</select>
			</span></b>
			<a href="javascript: void(0);" onclick="window.open('janelas_cadastros/rack.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=160'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a></td>
			<td colspan="2">&nbsp;<b><input type="text" name="cd_especialidade" id="cd_especialidade" size="3" class="inputs" value="<%=cd_especialidade%>" onKeyup="mostra_espec();" onkeypress="detecta_enter(event,'cd_procedimento_busca','<%=cd_user%>')">
				<input type="hidden<%if mostra_teste = 1 then response.write("s")%>" name="cd_especialidade_x" id="cd_especialidade_x" size="3" value="<%=cd_especialidade%>">
			<span id='divEspec'><select name="cd_especialidade_1" id="cd_especialidade_1" class="inputs" onkeypress="detecta_enter(event,'cd_procedimento_busca','<%=cd_user%>')">
								<option value=""></option>
								<%strsql ="Select * from TBL_espec where a057_status = 1 order by nm_especialidade"
						  		Set rs_espec = dbconn.execute(strsql)%>
								<%Do while Not rs_espec.EOF%><%cd_espec = rs_espec("cd_codigo")%>
								<%if cd_especialidade=cd_espec Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
								<option value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("cd_codigo")&" - "&rs_espec("nm_especialidade")%></option><%rs_espec.movenext
						 		Loop
								
								rs_espec.close
								set rs_espec = nothing%>
							</select>
			</span>
			<a href="javascript: void(0);" onclick="window.open('janelas_cadastros/especialidade.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=160'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a></td>
			<%if cd_co = 1 then
				co_ck = "checked"
			end if%>
			<!--td>&nbsp;</td-->
		</tr>
	
		<tr bgcolor="#f7f7f7"><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0" style="">
			<td colspan="100%">&nbsp;Procedimento</td>
		</tr>
		<tr>
			<td colspan="100%" bgcolor="#f7f7f7">
			<b><input type="text" name="cd_procedimento_busca" id="cd_procedimento_busca" size="10" maxlength="30" class="inputs"  onKeyup="mostra_procedimento();" onkeypress="detecta_enter(event,'somar_proc','<%=cd_user%>')">
			&nbsp;<span id='divPro_mostra'>&nbsp;<input type="Button" name="somar" value="ok"></span>
			<!--input type="hiddens" name="cd_procedimento_2" value="">proc2-->
			<input type="hidden" name="procedimentos_total" id="procedimentos_total" value="<%=proc_altera%>">
			<input type="hidden" name="proc_res" size="5">
			<span id='divproc_lista'>&nbsp;</span>
			<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/procedimento.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=190'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
			</td>
		</tr>
		<tr>
			<td colspan="100%" bgcolor="#f5f5f5"><%'***** Substitui a vírgula pela sentença de busca *****
		'if not no_protocolo = "" Then
		if not cd_cod_protocolo = "" Then
			strsql ="Select * from TBL_protocolo_procedimento where cd_protocolo="&cd_cod_protocolo&""
			Set rs_proc = dbconn.execute(strsql)
			
			%><table border="0">
				<tr>
					<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td align="right" valign="top" style="width:100px; color:white;" bgcolor="#d1d1d1">Cód. AMB&nbsp;</td>
					<td style="width:450px" style="width:100px; color:white;" bgcolor="#d1d1d1">&nbsp;Procedimento</td>
					<td bgcolor="#d1d1d1"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
				</tr>
			<%num_cor = 1
			Do while not rs_proc.EOF
			cd_procedimento = rs_proc("a058_codpro")
			cd_protocolo = rs_proc("cd_protocolo")
				cd_procedimento = replace(cd_procedimento,","," OR a058_codpro=")%>
			
				<%strsql ="Select * from TBL_procedimento where cd_codigo="&cd_procedimento&""
				Set rs = dbconn.execute(strsql)%>
				<%Do while not rs.EOF
				if num_cor = 1 then cor_linha = "#e5e5e5" else cor_linha = "#ffffff" end if%>
					<tr>
						<td bgcolor="#f5f5f5" ><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
						<td align="right" valign="top" style="width:100px" bgcolor="<%=cor_linha%>"><%'=rs("cd_codigo")%>
						<%=mid(rs("cd_codigo"),1,2)&"."&mid(rs("cd_codigo"),3,2)&"."&mid(rs("cd_codigo"),5,3)&"."&mid(rs("cd_codigo"),8,1)&"<!--br-->"%>&nbsp;</td>
						<td style="width:450px" bgcolor="<%=cor_linha%>">&nbsp;<%=left(rs("nm_procedimento"),80)&"<br>"%> <%'=cd_protocolo%></td>
						<td bgcolor="<%=cor_linha%>"><img src="imagens/ic_del.gif" onclick="javascript:JsProcDelete('<%=rs_proc("cd_codigo")%>','<%=no_protocolo%>','<%=cd_unidade%>','<%=no_digito%>')" type="button" value="A" title="Apagar"></td>
					</tr>				
				<%if num_cor < 2 then
					num_cor = num_cor + 1
				else
					num_cor = 1
				end if
				
				rs.movenext
				loop
				
				rs.close
				set rs = nothing%>								
			<%rs_proc.movenext
		loop
			
		rs_proc.close
		set rs_proc = nothing%></table>
		
		<%end if%>
			</b><br></td>
		</tr>
		
		
<!--tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%" style="color:white;"><b>&nbsp;Emprestimos</b></td>
		</tr>
		<tr bgcolor="#f7f7f7">
			<td colspan="100%" style="color:gray;">
			&nbsp; &nbsp; &nbsp;Material &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; Qtd.<br>
			<b><!--input type="text" name="cd_material_busca" size="10" maxlength="30" class="inputs"  onKeyup="mostra_material();" onFocus="nextfield ='qtd_material';" onBlur="mostra_material();"-->
			<input type="text" name="cd_material_busca" id="cd_material_busca" size="10" maxlength="30" class="inputs"  onKeyup="mostra_material();" onBlur="mostra_material();" onkeypress="detecta_enter(event,'qtd_material','<%=cd_user%>')">
			<input type="text" name="qtd_material" id="qtd_material" size="5" maxlength="3" class="inputs"  onkeypress="detecta_enter(event,'somar_mat','<%=cd_user%>')">
			<span id='divMat_mostra'>&nbsp;</span><!-- resultado do campo cd_material --><br>
			<!--input type="hidden" name="cd_material_2" value=""-->
			<input type="hidden" name="materiais_total" id="materiais_total" value="">
			<input type="hidden" name="res_mat" size="5">			
			<span id='divMat_lista'><!-- resultado da soma-->
			<%mat_array = split(materiais,",")
		For Each mat_item In mat_array
			
			qtd_carac = len(mat_item)
			separador = Instr(mat_item,";")
			'response.write(mat_item)
		
				xsql = "SELECT * FROM TBL_material where A058_codmat="&mat_item&" order by A058_desmat"
				Set rs_mat1 = dbconn.execute(xsql)%>
				<textarea cols="30" rows="5" name="materiais"><%'=cd_material%>
				<%while not rs_mat1.EOF
				response.write(rs_mat1("A058_codmat")&" - "&rs_mat1("A058_desmat")&"<br>")
				rs_mat1.MoveNext
				 wend
				
				  rs_mat1.close
				  set rs_mat1 = nothing%>			  
				  </textarea>
				  <input type="hidden" name="materiais" value="">
				  
		<%Next%>
			</span>
			<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/material.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=170'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->			 
			</td>		
		</tr>
		<tr bgcolor="#f7f7f7">
			<td colspan="100%" bgcolor="#f5f5f5"><%'***** Substitui a vírgula pela sentença de busca *****
		'if not no_protocolo <> "" Then
		if cd_cod_protocolo <> "" Then
			strsql ="Select * from TBL_protocolo_material where cd_protocolo="&cd_cod_protocolo&""
			Set rs_mat = dbconn.execute(strsql)%>
			<table border="0">
				<tr>
					<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<!--td align="right" valign="top" style="width:100px; color:white;" bgcolor="#d1d1d1">Código&nbsp;</td-->
					<td align="right" valign="top" style="width:100px; color:white;" bgcolor="#d1d1d1">Qtd&nbsp;</td>
					<td style="width:450px" style="width:100px; color:white;" bgcolor="#d1d1d1">&nbsp;Item</td>
					<td bgcolor="#d1d1d1"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
				</tr>
			<%num_cor = 1
			Do while not rs_mat.EOF
			cd_material = rs_mat("a061_codmat")
				cd_material = replace(cd_material,","," OR a061_codmat=")%>
			
				<%strsql ="Select * from TBL_material where a061_codmat="&cd_material&""
				Set rs_material = dbconn.execute(strsql)%>
				<%Do while not rs_material.EOF%>
					<%if num_cor = 1 then cor_linha = "#f5f5f5" else cor_linha = "#ffffff" end if%>
					<tr>
						<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
						<td align="right" valign="top" bgcolor="<%=cor_linha%>"><%'=rs_material("a061_codmat")%><%=rs_mat("a010_quantidade")%>&nbsp;-</td>
						<td bgcolor="<%=cor_linha%>">&nbsp;<%=left(rs_material("a061_nommat"),80)&"<br>"%></td>
						<td bgcolor="<%=cor_linha%>"><img src="imagens/ic_del.gif" onclick="javascript:JsMatDelete('<%=rs_mat("cd_codigo")%>','<%=no_protocolo%>','<%=cd_unidade%>','<%=no_digito%>')" type="button" value="A" title="Apagar"></td>
					</tr>				
				<%if num_cor < 2 then
					num_cor = num_cor + 1
				else
					num_cor = 1
				end if
				
				rs_material.movenext
				loop
				
				rs_material.close
				set rs_material = nothing%>	
											
			<%rs_mat.movenext
		loop
			
			rs_mat.close
			set rs_mat = nothing%></table>
		
		<%end if%>
			</b><br></td>
		</tr>	
		<tr bgcolor="#f7f7f7"><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>		
<%if oticas = xx Then%>
		<!-- **********************************************************************
		***************************************************************************
		***** NOVA ÁREA DE INCLUSÃO DE ÓTICAS*****-->		
		
		<tr bgcolor="#808080">
			<td align="center" colspan="100%" style="color:white;"><b>&nbsp;Óticas</b></td>
		</tr>
		<tr>
			<td colspan="100%">Nº Patrimônio</td>
		</tr>
		<tr>
			<td colspan="100%" bgcolor="#f7f7f7">
			<b><input type="text" name="cd_patrimonio_busca" id="cd_patrimonio_busca" size="10" maxlength="30" class="inputs" onKeyup="setTimeout('mostra_patrimonio()',250);" onkeypress="detecta_enter(event,'somar_pat','<%=cd_user%>')">
			<input type="hidden" name="patrimonio_busca">
			&nbsp; <span id='divPat_mostra'>&nbsp;
				<input type="Button" name="somar_pat" id="somar_pat" value="ok">
			</span>
			<br>			
			<input type="hiddens" name="subtotal_patrimonio" id="subtotal_patrimonio" value=""><br>
			<span id='divPat_lista'> &nbsp;</span>
			
			<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/procedimento.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=190'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a--></td>
		</tr>
		<tr bgcolor="#f7f7f7">
			<td colspan="100%" bgcolor="#f5f5f5"><%'***** Substitui a vírgula pela sentença de busca *****
		if cd_cod_protocolo <> "" Then
			strsql ="Select * from VIEW_protocolo_patrimonio_lista where cd_protocolo="&cd_cod_protocolo&" order by no_patrimonio"
			Set rs_pat = dbconn.execute(strsql)
			if not rs_pat.EOF Then 
				no_patrimonio = rs_pat("no_patrimonio")
			end if%>
			<table border="0">
				<tr>
					<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					<td align="right" valign="top" style="width:100px; color:white;" bgcolor="#d1d1d1">Patrimonio</td>
					<td style="width:450px" style="width:100px; color:white;" bgcolor="#d1d1d1">&nbsp;Item</td>
					<td bgcolor="#d1d1d1"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
				</tr>
			<%num_cor = 1
			Do while not rs_pat.EOF%>
				
				<%if num_cor = 1 then cor_linha = "#e5e5e5" else cor_linha = "#ffffff" end if%>
					<tr>
						<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"><%=cd_pat%></td>
						<td align="right" valign="top" style="width:100px" bgcolor="<%=cor_linha%>"><%=rs_pat("nm_tipo")%><%=rs_pat("no_patrimonio")%>&nbsp;</td>
						<td style="width:450px" bgcolor="<%=cor_linha%>">&nbsp; <%=left(rs_pat("nm_equipamento"),80)&"<br>"%></td>
						<td bgcolor="<%=cor_linha%>"><img src="imagens/ic_del.gif" onclick="javascript:JsPatDelete('<%=rs_pat("cd_codigo")%>','<%=no_protocolo%>','<%=cd_unidade%>','<%=no_digito%>')" type="button" value="A" title="Apagar"></td>
					</tr>				
				<%if num_cor < 2 then
					num_cor = num_cor + 1
				else
					num_cor = 1
				end if
				
				
				%>								
			<%rs_pat.movenext
		loop
			
			
			rs_pat.close
			set rs_pat = nothing%></table>
		<%end if%>
			</b><br></td>
		</tr>
		<!--******************************************************* Fim da inclusao das Oticas  ***********************************************************************-->
<%if asd = 1 then%>
<%end if

end if%>
<SCRIPT LANGUAGE="javascript">
	function JsProcDelete(cod1,cod2,cod3,cod4)
	{
	  if (confirm ("Tem certeza que deseja excluir o procedimento?"))
		  {
			document.location=('protocolo/acoes/protocolo_acao.asp?cd_codigo='+cod1+'&no_protocolo='+cod2+'&cd_unidade='+cod3+'&no_digito='+cod4+'&acao=excluir&obj_excl=6');
		  }
	}
	
	function JsMatDelete(cod1,cod2,cod3,cod4)
	{
	  if (confirm ("Tem certeza que deseja excluir o material de empréstimo?"))
		  {
			document.location=('protocolo/acoes/protocolo_acao.asp?cd_material='+cod1+'&no_protocolo='+cod2+'&cd_unidade='+cod3+'&no_digito='+cod4+'&obj_excl=3&acao=excluir');
		  }
	}
	
	function JsPatDelete(cod1,cod2,cod3,cod4)
	{
	  if (confirm ("Tem certeza que deseja excluir esta ótica?"))
		  {
			document.location=('protocolo/acoes/protocolo_acao.asp?cd_codigo='+cod1+'&no_protocolo='+cod2+'&cd_unidade='+cod3+'&no_digito='+cod4+'&obj_excl=7&acao=excluir');
		  }
	}

</SCRIPT>
<!--SCRIPT LANGUAGE="javascript">
		 {
		   	MaskInput(document.Form.cd_idade, "999");			
			MaskInput(document.Form.nm_registro, "99999999999999");			
			MaskInput(document.Form.dt_mes_cirurgia, "99");
			MaskInput(document.Form.dt_ano_cirurgia, "9999");			
			MaskInput(document.Form.dt_hora_agenda, "99");
			MaskInput(document.Form.dt_minuto_agenda, "99");			
			MaskInput(document.Form.dt_hora_inicio, "99");
			MaskInput(document.Form.dt_minuto_inicio, "99");			
			MaskInput(document.Form.dt_hora_fim, "99");
			MaskInput(document.Form.dt_minuto_fim, "99");		
		 }
		</SCRIPT-->