		<tr bgcolor="#f5f5f5">
			<td colspan="100%">&nbsp;Protocolo: <b><%=zero(int(cd_unidade))%>.<%=proto(cd_protocolo)%>-<%=dv%> &nbsp;<%=mensagem%></b></td>
		</tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font color="white"><b>&nbsp;Identifica��o do Paciente - <%=acao%></b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td align="left" colspan="3">&nbsp;Nome</td>
			<td align="left" colspan="2">&nbsp;Idade</td>
		</tr>
		<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">.
		<input type="hidden" name="cd_protocolo"  value="<%=cd_protocolo%>">
		<input type="hidden" name="cd_digito" value="<%=cd_digito%>">
		<input type="hidden" name="acao" value="inserir">
		&nbsp;				
			
		<tr>
			<td align="left" colspan="3">&nbsp;<b><input type="text" name="nm_nome" size="60" maxlength="60" class="inputs" value="<%=nm_nome%>" onFocus="nextfield ='cd_idade';"></b></td>
			<td align="left">&nbsp;<b><input type="text" name="cd_idade" size="5" maxlength="3" class="inputs" value="<%=cd_idade%>" onFocus="nextfield ='nm_registro';"></b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td align="left">&nbsp;Registro</td>
			<td align="left" colspan="2">&nbsp;Conv�nio: </td>
			<td align="left">&nbsp;Sexo: </td>
		</tr>
		<tr>
			<td align="left">&nbsp;<b><input type="text" name="nm_registro" size="10" maxlength="15" class="inputs" value="<%=cd_registro%>" onFocus="nextfield ='cd_convenio';"></b>&nbsp;&nbsp;
			<td align="left" valign="top" colspan="2">&nbsp;
			<b>					<input type="text" name="cd_convenio" size="3" maxlength="4" class="inputs" value="<%=cd_convenio%>" onFocus="nextfield ='cd_convenio_1';">	
			<span id='divAssist_m1'><select name="cd_convenio_1" class="inputs" onFocus="nextfield ='nm_sexo';"><option value="">
								<%strsql ="Select * from TBL_convenios order by A054_descon"
						  		Set rs_conv = dbconn.execute(strsql)%><%Do while Not rs_conv.EOF%><%cd_conv = rs_conv("A054_codcon")%>
								<%if cd_convenio=cd_conv Then%><%ck_conv="selected"%><%else ck_conv=""%><%end if%>
								<option value="<%=rs_conv("A054_codcon")%>" <%=ck_conv%>><%=rs_conv("A054_descon")%></option><%rs_conv.movenext
						 		Loop%>
							</select>
			</span>
			<a href="javascript: void(0);" onclick="window.open('janelas_cadastros/assist_m_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=160'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a>	
			</td>
			<td align="left">&nbsp;<b><select name="nm_sexo" size="1" class="inputs" onFocus="nextfield ='nm_cirug_realizada';">
							 <option value=""></option>
							 <%if nm_sexo="M" Then%><%ck_m = "selected"%>
							 <%Elseif nm_sexo="F" Then%><%ck_f = "selected"%>
							 <%end if%>
							 <option value="M" <%=ck_m%>>M</option>
							 <option value="F" <%=ck_f%>>F</option>&nbsp;&nbsp;
							 </select>
							 </b></td>					
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha 1--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font color="white"><b>&nbsp;Dados Gerais</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Cirurgia Realizada?</td>
			<td colspan="2">&nbsp;Data</td>
			<td>&nbsp;Agenda</td>
		</tr>
		<tr>
			<td>&nbsp;<b><select name="nm_cirug_realizada" class="inputs" onFocus="nextfield ='dt_dia_cirurgia';">
					<option value=""></option>
					<%if nm_cirug_realizada="S" Then%><%ck_real_s = "selected"%>
					<%Elseif nm_cirug_realizada="N" Then%><%ck_real_n = "selected"%>
					<%end if%>
					<option value="S" <%=ck_real_s%>>SIM</option>
					<option value="N" <%=ck_real_n%>>N�O</option>
				</select></b></td>
					<%if dt_procedimento <> "" Then
						dt_proced_dia = DAY(dt_procedimento)
						dt_proced_mes = MONTH(dt_procedimento)
						dt_proced_ano = YEAR(dt_procedimento)
					end if%>
			<td colspan="2"><b><input type="text" name="dt_dia_cirurgia" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=zero(dt_proced_dia)%>" onFocus="nextfield ='nm_agenda_cirurgia';">
							<input type="text" name="dt_mes_cirurgia" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=zero(dt_proced_mes)%>" onFocus="nextfield ='nm_agenda_cirurgia';">
							<input type="text" name="dt_ano_cirurgia" size="4" maxlength="4" maxlength="3" class="inputs" value="<%=dt_proced_ano%>" onFocus="nextfield ='nm_agenda_cirurgia';"></b></td>
							<!--<input type="text" name="dt_ano_cirurgia" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_proced_ano%>" onFocus="nextfield ='nm_agenda_cirurgia';"></b></td-->
			<td>&nbsp;<b>
				<select name="nm_agenda_cirurgia" class="inputs" onFocus="nextfield ='dt_hora_agenda';">
					<option value=""></option>
					<%if nm_agenda="A" Then%><%ck_agd_a = "selected"%>
					<%Elseif nm_agenda="E" Then%><%ck_agd_e = "selected"%>
					<%Elseif nm_agenda="N" Then%><%ck_agd_n = "selected"%>
					<%Elseif nm_agenda="P" Then%><%ck_agd_p = "selected"%>
					<%Elseif nm_agenda="U" Then%><%ck_agd_u = "selected"%>
					<%end if%>
					<option value="N" <%=ck_agd_n%>>Normal</option>
					<option value="A" <%=ck_agd_a%>>A seguir</option>
					<option value="E" <%=ck_agd_e%>>Empr�stimo</option>					
					<option value="P" <%=ck_agd_p%>>P�s Agendada</option>
					<option value="U" <%=ck_agd_u%>>Urg�ncia</option>					
				</select>
				</b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Hora Agendada</td>
			<td colspan="2">&nbsp;In�cio</td>
			<td>&nbsp;T�rmino</td>
		</tr>
		<%
		if dt_hr_agenda <> "" Then
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
		end if
		%>
		<tr>
			<td>&nbsp;<b><input type="text" size="2" name="dt_hora_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_hr_agenda)%>" onFocus="nextfield ='dt_hora_inicio';"> :
				<input type="text" size="2" name="dt_minuto_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_min_agenda)%>" onFocus="nextfield ='dt_hora_inicio';"></b></td>
			<td colspan="2">&nbsp;<b><input type="text" size="2" name="dt_hora_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_hr_inicio)%>" onFocus="nextfield ='dt_hora_fim';"> :
				<input type="text" size="2" name="dt_minuto_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_min_inicio)%>" onFocus="nextfield ='dt_hora_fim';"></b></td>
			<td>&nbsp;<b><input type="text" size="2" name="dt_hora_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_hr_fim)%>" onFocus="nextfield ='cd_crm';"> :
				<input type="text" size="2" name="dt_minuto_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="2" class="inputs" value="<%=zero(dt_min_fim)%>" onFocus="nextfield ='cd_crm';"></b></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td align="center" colspan="100%"><font color="white"><b>&nbsp;Procedimento Cir�rgico</b></font></td>
		</tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;CRM &nbsp; / M�dico</td>
		</tr>
		<tr>
			<td colspan="100%">&nbsp;<b><input type="text" name="cd_crm" class="inputs" size="10" maxlength="10" value="<%=cd_med%>" onFocus="nextfield ='cd_crm_1';">
			<%strsql ="Select * from TBL_medicos order by A055_codmed"
						  		Set rs_med = dbconn.execute(strsql)
								if not rs_med.EOF Then
								cd_crm = rs_med("A055_crmmed")
								end if%>
							<select name="cd_crm_1" class="inputs" onFocus="nextfield ='cd_rack';" onFocus="nextfield ='cd_rack';">
								<option value=""></option>
								<%Do while Not rs_med.EOF%>
								<%cd_medico = rs_med("A055_codmed")%>
								<%if int(cd_medico)=int(cd_med) Then%>
								<%ck_med="selected"%><%else ck_med=""%>
								<%end if%>
								<option value="<%=rs_med("A055_codmed")%>" <%=ck_med%>><%=cd_medico%> - <%=rs_med("A055_desmed")%></option><%rs_med.movenext
						 		Loop%>
							</select>
			</b></td>							
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td>&nbsp;Rack</td>
			<td colspan="3">&nbsp;Especialidade</td>
		</tr>
		<tr>
			<td>&nbsp;<b><input type="text" name="cd_rack" size="3" class="inputs" value="<%=cd_rack%>" onFocus="nextfield ='cd_rack_1';">
							<select name="cd_rack_1" class="inputs" onFocus="nextfield ='cd_especialidade';">
								<option value=""></option>
								<%strsql ="Select * from TBL_Rack"' order by nm_rack"
						  		Set rs_rack = dbconn.execute(strsql)%>
								<%Do while Not rs_rack.EOF%><%cd_rk = rs_rack("A056_codrac")%>
								<%if cd_rack=int(cd_rk) Then%><%ck_rack="selected"%><%else ck_rack=""%><%end if%>
								<option value="<%=rs_rack("A056_codrac")%>" <%=ck_rack%>><%=rs_rack("A056_desrac")%></option><%rs_rack.movenext
						 		Loop%>
							</select></b></td>							
			<td>&nbsp;<b><input type="text" name="cd_especialidade" size="3" class="inputs" value="<%=cd_especialidade%>" onFocus="nextfield ='cd_especialidade_1';">
							<select name="cd_especialidade_1" class="inputs" onFocus="nextfield ='cd_procedimento';">
								<option value=""></option>
								<%strsql ="Select * from TBL_espec order by nm_especialidade"
						  		Set rs_espec = dbconn.execute(strsql)%>
								<%Do while Not rs_espec.EOF%><%cd_espec = rs_espec("cd_codigo")%>
								<%if cd_especialidade=cd_espec Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
								<option value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option><%rs_espec.movenext
						 		Loop%>
							</select></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#e0e0e0">
			<td colspan="100%">&nbsp;Procedimento</td>
		</tr>
		<tr>
			<td colspan="100%">
			<b><input type="text" name="cd_procedimento" size="10" maxlength="10" class="inputs" onFocus="nextfield ='somar';">
			<input type="button" name="somar" value="+" onclick="soma(document.form.cd_procedimento.value,document.form.cd_procedimento_1.value,document.form.procedimentos.value)" onFocus="nextfield ='cd_procedimento';">
			&nbsp; <select name="cd_procedimento_1" class="inputs"  onclick="soma(document.form.cd_procedimento.value,document.form.cd_procedimento_1.value,document.form.procedimentos.value)" onFocus="nextfield ='done';">
				<option value=""></option>
				<%strsql ="Select * from TBL_procedimento order by A058_despro"
		  		Set rs_proc = dbconn.execute(strsql)%>
				<%Do while Not rs_proc.EOF%><%cd_proc = rs_proc("A058_codpro")%>
				<option value="<%=rs_proc("A058_codpro")%>" <%'=ck_proc%>>(<%=rs_proc("A058_codpro")%>) - <%=left(rs_proc("A058_despro"),25)%></option><%rs_proc.movenext
		 		Loop%>
			</select>
			
			<!--textarea cols="30" rows="5" name="procedimentos"><%=cd_procedimento%></textarea-->
			<input type="hidden" name="procedimentos" value="<%=cd_procedimento%>">			
			<input type="hidden" name="res" size="5">
			<img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_prot();"><br>
			<br>
			<span id='divprot'>
			<%proc_array = split(procedimentos,",")
		For Each proc_item In proc_array
		
			xsql = "SELECT * FROM TBL_procedimento where A058_codpro="&proc_item&" order by A058_despro"
			Set rs_proc1 = dbconn.execute(xsql)%>
			<textarea cols="30" rows="5" name="procedimentos"><%'=cd_procedimento%>
			<%while not rs_proc1.EOF
			response.write(rs_proc1("A058_codpro")&" - "&rs_proc1("A058_despro")&"<br>")
			rs_proc1.MoveNext
			 wend
			
			  rs_proc1.close
			  set rs_proc1 = nothing%>			  
			  </textarea>
			  <%
		
		'numero = numero + 1
		
		Next%>
			</span>
						 
			</td>		
		</tr>
		<tr>
			<td colspan="100%" bgcolor="#f5f5f5"><%'***** Substitui a v�rgula pela senten�a de busca *****
		if not protocolo = "" Then
			strsql ="Select * from TBL_protocolo_procedimento where A002_numseq="&protocolo&""
			Set rs_proc = dbconn.execute(strsql)
			
			%><table border="0"><%
			Do while not rs_proc.EOF
			cd_procedimento = rs_proc("a058_codpro")'cd_codigo")
				cd_procedimento = replace(cd_procedimento,","," OR a058_codpro=")%>
			
				<%strsql ="Select * from TBL_procedimento where a058_codpro="&cd_procedimento&""
				Set rs = dbconn.execute(strsql)%>				
				<%Do while not rs.EOF%>
					<tr>
						<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
						<td align="right" valign="top"><%=rs("a058_codpro")%>&nbsp;-</td>
						<td><%=rs("A058_despro")&"<br>"%></td>
						<td bgcolor="#f5f5f5"><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
					</tr>				
				<%rs.movenext
				loop%>								
			<%rs_proc.movenext
			loop
			%></table><%
			end if%>
			</b><br></td>
		</tr>
		<tr>
			<td><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>
			<td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="270" height="1" border="0"></td>
			<td><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td>							
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>		
		
		
		
		
		
		
		
		
		
		
		
		