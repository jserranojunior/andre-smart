<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->


<%cd_situacao = request("mostra_movimentacao")
response.write(cd_situacao&"<br><br>")


if x = "" then

limite_sit = instr(1,cd_situacao,":",1)
limite_sit2 = instr(1,cd_situacao,"{",1)
limite_sit3 = instr(1,cd_situacao,"}",1)
limite_sit4 = instr(1,cd_situacao,"[",1)

str_situacao = mid(cd_situacao,1,limite_sit)
	str_situacao = replace(str_situacao,":","")
cd_unidade = mid(cd_situacao,limite_sit,limite_sit2-limite_sit)
	cd_unidade = replace(cd_unidade,":","")
cd_rack = mid(cd_situacao,limite_sit2,limite_sit3-limite_sit2)
	cd_rack = replace(cd_rack,"{","")

cd_unidade2 = mid(cd_situacao,limite_sit3,limite_sit4-limite_sit3)
	cd_unidade2 = replace(cd_unidade2,"}","")
cd_rack2 = mid(cd_situacao,limite_sit4,len(cd_situacao))
	cd_rack2 = replace(cd_rack2,"[","")%>
	
<table border="0" class="txt" class="no_print" style="font-size:12px; font-family:arial; border-collapse:collapse;" align="left">	
			<!--form id="form_ajax" name="form_ajax"-->
<%if str_situacao = "" Then

elseif str_situacao = 0 then
'*** - ***%>
<tr>
	<td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><input type="hidden" name="cd_unidade_mov" value="<%=cd_unidade%>"></td>
	<td>&nbsp; Data: <input type="text" name="dia_empr_devolucao" size="2" maxlength="2" value="<%=dia_devolucao%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="mes_empr_devolucao" size="2" maxlength="2" value="<%=mes_devolucao%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="ano_empr_devolucao" size="4" maxlength="4" value="<%=ano_devolucao%>"></td>
</tr>

<%elseif str_situacao = 1 then
'*** Descarte ***
cor_emprest = "#ffdfdf"%>
<tr style="background-color:<%=cor_emprest%>;">
	<td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><input type="hidden" name="cd_unidade_mov" value="<%=cd_unidade%>"><%'="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%></td>
	<td>&nbsp; Data: <input type="text" name="dia_movimento" size="2" maxlength="2" value="<%'=dia_devolucao%>" onKeyup="mov_descarte();" id="dia_movimento">/<input type="text" name="mes_movimento" size="2" maxlength="2" value="<%'=mes_devolucao%>" onKeyup="mov_descarte();" id="mes_movimento">/<input type="text" name="ano_movimento" size="4" maxlength="4" value="<%'=ano_devolucao%>" onKeyup="mov_descarte();" id="ano_movimento"><br>
		<img src="../../imagens/px.gif" alt="" width="215" height="1" border="0"></td>
	<td>&nbsp;Obs.:<br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
	<td><textarea cols="25" rows="1" name="nm_observacoes" onKeyup="mov_descarte();" id="nm_observacoes"><%'=nm_empr_obs%></textarea>
		<img src="../../imagens/px.gif" alt="" width="60" height="1" border="0"></td>
</tr>
<%elseif str_situacao = 2 then
'*** Manutencao ***
	cor_emprest = "#d9e8e8"%>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp; Data<br><img src="../../imagens/blackdot.gif" alt="" width="150" height="1" border="0"></td>
		<td colspan="3"><input type="text" name="dia_emprestimo" size="2" maxlength="2" value="<%'=dia_emprestimo%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="mes_emprestimo" size="2" maxlength="2" value="<%'=mes_emprestimo%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="ano_emprestimo" size="4" maxlength="4" value="<%'=ano_emprestimo%>"  onFocus="nextfield ='cd_unidade_des';"></td>
	</tr>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>Unidade destino</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_hospital >= 1 and cd_status=1"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade_des" class="inputs" onFocus="nextfield ='cd_rack_des';">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						str_unidade_des = rs_uni("cd_codigo")%>
						<option value="<%=rs_uni("cd_codigo")%>" <%'if int(cd_unidade_des)=str_unidade_des then response.write("SELECTED") End if%>><%=rs_uni("nm_unidade")%></option>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
						</select>					
		</td>
		<td> &nbsp; Rack destino:</td>
		<%strsql ="SELECT * FROM TBL_rack order by A056_desrac"
		Set rs_rack = dbconn.execute(strsql)%>
		<td>
		<select name="cd_rack_des" class="inputs" onFocus="nextfield ='nm_solicitante';">
			<option value="" SELECTED>&nbsp;</option>
			<%while not rs_rack.EOF%>
				<option value="<%=rs_rack("A056_codrac")%>" <%'if int(cd_rack_des) = rs_rack("A056_codrac") then response.write("SELECTED") End if%>><%=rs_rack("A056_desrac")%></option>
			<%rs_rack.movenext
			wend%>
		</select></td>
	</tr>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>Solicitante:</td>
		<td><input type="text" name="nm_solicitante" value="<%'=nm_solicitante%>" size="34" maxlength="200" onFocus="nextfield ='nm_empr_obs';"></td>
		<td> &nbsp; Obs.:<br><img src="../../imagens/blackdot.gif" alt="" width="150" height="1" border="0"></td>
		<td><textarea cols="25" rows="1" name="nm_empr_obs"><%'=nm_empr_obs%></textarea></td>				
	</tr>
<%elseif str_situacao = 3 then
'*** OK ***%>
<tr>
	<td>
		OK
	</td>
</tr>
<%elseif str_situacao = 4 then
'*** Emprestimo ***
cor_emprest = "#e1fbd2"%>
<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp;Data <br><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><%'="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%></td>
		<td><input type="text" name="dia_movimento" size="2" maxlength="2" value="<%'=dia_emprestimo%>" onblur="mov_emprestimo();" id="dia_movimento" onkeyup="javascript:JumpField(this,'mes_movimento');">/
		<input type="text" name="mes_emprestimo" size="2" maxlength="2" value="<%'=mes_emprestimo%>"  onblur="mov_emprestimo();" id="mes_movimento" onkeyup="javascript:JumpField(this,'ano_movimento');">/
		<input type="text" name="ano_emprestimo" size="4" maxlength="4" value="<%'=ano_emprestimo%>" onKeyup="mov_emprestimo();" id="ano_movimento"></td>
		
		<td>&nbsp;Unid. dest</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_hospital >= 1 and cd_status=1"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade_des" class="inputs" onFocus="nextfield ='cd_rack_des';" onChange="mov_emprestimo();" id="und_destino">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						str_unidade_des = rs_uni("cd_codigo")%>
							<%if abs(cd_unidade) <> abs(str_unidade_des) then%><option value="<%=rs_uni("cd_codigo")%>" <%'if int(cd_unidade_des)=str_unidade_des then response.write("SELECTED") End if%>><%=rs_uni("nm_unidade")%><%'=cd_unidade&"-"&str_unidade_des%></option><%end if%>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
						</select></td>
	</tr>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp;Solicitante:</td>
		<td><input type="text" name="nm_solicitante" size="30" maxlength="30" onKeyup="mov_emprestimo();" id="nm_solicitante"></td>
		<td>&nbsp;Obs.:<br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><textarea cols="25" rows="1" name="nm_observacoes" onKeyup="mov_emprestimo();" id="nm_observacoes"><%'=nm_empr_obs%></textarea></td>	
	</tr>
	
<%elseif str_situacao = 5 then
'*** Comodato ***
cor_emprest = "#ccccff"%>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp;Data<br><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><%'="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%></td>
		<td><input type="text" name="dia_movimento" size="2" maxlength="2" value="<%'=dia_emprestimo%>" onBlur="mov_comodato();" id="dia_movimento" onkeyup="javascript:JumpField(this,'mes_movimento');">/
			<input type="text" name="mes_movimento" size="2" maxlength="2" value="<%'=mes_emprestimo%>" onBlur="mov_comodato();" id="mes_movimento" onkeyup="javascript:JumpField(this,'ano_movimento');">/
			<input type="text" name="ano_movimento" size="4" maxlength="4" value="<%'=ano_emprestimo%>" onKeyup="mov_comodato();" id="ano_movimento"></td>
		<td>&nbsp;Obs.:<br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><textarea cols="25" rows="1" name="nm_observacoes" onKeyup="mov_comodato();" id="nm_observacoes"><%'=nm_empr_obs%></textarea></td>
	</tr>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp;Unidade destino</td>
			<%strsql = "SELECT * FROM TBL_unidades where cd_hospital >= 1 and cd_status=1"
			Set rs_uni = dbconn.execute(strsql)%>
	    <td><select name="cd_unidade_des" class="inputs" onChange="mov_comodato();" id="und_destino">
						<option value=""></option>
						<%Do While Not rs_uni.eof
						str_unidade_des = rs_uni("cd_codigo")%>
							<%if abs(str_unidade_des) <> abs(cd_unidade2) then%><option value="<%=rs_uni("cd_codigo")%>"><%=rs_uni("nm_unidade")%><%'=str_unidade_des&"-"&cd_unidade2%></option><%end if%>
						<%rs_uni.movenext
						unidade_check = ""
						loop%>
						</select></td>
		<td> &nbsp; Rack destino:</td>
		<%strsql ="SELECT * FROM TBL_rack order by A056_desrac"
		Set rs_rack = dbconn.execute(strsql)%>
		<td>
		<select name="cd_rack_des" class="inputs" onChange="mov_comodato();" id="rack_destino">
			<option value="" SELECTED>&nbsp;</option>
			<%while not rs_rack.EOF%>
				<option value="<%=rs_rack("A056_codrac")%>" <%'if int(cd_rack_des) = rs_rack("A056_codrac") then response.write("SELECTED") End if%>><%=rs_rack("A056_desrac")%></option>
			<%rs_rack.movenext
			wend%>
		</select></td>
	</tr>
	
<%elseif str_situacao = 6 then
'*** Estoque ***
cor_emprest = "#ffe3c6"%>
<tr style="background-color:<%=cor_emprest%>;">
	<td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><input type="hidden" name="cd_unidade_mov" value="<%=cd_unidade%>"><%'="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%></td>
	<td>&nbsp; Data: <input type="text" name="dia_movimento" size="2" maxlength="2" value="<%'=dia_devolucao%>" onKeyup="mov_estoque();" id="dia_movimento">/<input type="text" name="mes_movimento" size="2" maxlength="2" value="<%'=mes_devolucao%>" onKeyup="mov_rack();" id="mes_movimento">/<input type="text" name="ano_movimento" size="4" maxlength="4" value="<%'=ano_devolucao%>" onKeyup="mov_rack();" id="ano_movimento">
	<img src="../../imagens/px.gif" alt="" width="370" height="1" border="0"></td>
</tr>

<%elseif str_situacao = 7 then
'*** Rack ***
cor_emprest = "#fde6c6"%>
<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp;Data<img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><%'="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%>
		<input type="hidden" name="cd_rack_ori" value="<%=cd_rack%>">
		<input type="hidden" name="cd_unidade_des" value="<%=cd_unidade%>"></td>
		
		<td><input type="text" name="dia_movimento" size="2" maxlength="2" value="<%'=dia_emprestimo%>"  onblur="mov_rack();" id="dia_movimento" onkeyup="javascript:JumpField(this,'mes_movimento');">/
		<input type="text" name="mes_movimento" size="2" maxlength="2" value="<%'=mes_emprestimo%>"  onblur="mov_rack();" id="mes_movimento" onkeyup="javascript:JumpField(this,'ano_movimento');">/
		<input type="text" name="ano_movimento" size="4" maxlength="4" value="<%'=ano_emprestimo%>"   onKeyup="mov_rack();" id="ano_movimento"></td>
		
		<td> &nbsp; Rack dest.:<%=cd_rack2%></td>
		<%strsql ="SELECT * FROM TBL_rack order by A056_desrac"
		Set rs_rack = dbconn.execute(strsql)%>
		<td>
		<select name="cd_rack_des" class="inputs"  onChange="mov_rack();" id="rack_destino">
			<option value="" SELECTED>&nbsp;</option>
			<%while not rs_rack.EOF%>
				<option value="<%=rs_rack("A056_codrac")%>" <%'if int(cd_rack_des) = rs_rack("A056_codrac") then response.write("SELECTED") End if%>><%=rs_rack("A056_desrac")%></option>
			<%rs_rack.movenext
			wend%>
		</select></td>
	</tr>
	<tr style="background-color:<%=cor_emprest%>;">
		<td>&nbsp;Solicitante:</td>
		<td><input type="text" name="nm_solicitante" value="<%'=nm_solicitante%>" size="34" maxlength="200" onKeyup="mov_rack();" id="nm_solicitante"></td>
		<td>&nbsp;Obs.:<br><img src="../../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
		<td><textarea cols="22" rows="1" name="nm_observacoes" onKeyup="mov_rack();" id="nm_observacoes"><%'=nm_empr_obs%></textarea></td>	
	</tr>
<%elseif str_situacao = 8 then
'*** Devolução ***%>
<tr>
	<td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><input type="hidden" name="cd_unidade_mov" value="<%=cd_unidade%>"><br><%="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%></td>
	<td>&nbsp; Data: <input type="text" name="dia_empr_devolucao" size="2" maxlength="2" value="<%=dia_devolucao%>" onblur="mov_devol();" id="dia_movimento"  onkeyup="javascript:JumpField(this,'mes_movimento');" >/
				 <input type="text" name="mes_empr_devolucao" size="2" maxlength="2" value="<%=mes_devolucao%>" onblur="mov_devol();" id="mes_movimento" onkeyup="javascript:JumpField(this,'ano_movimento');">/
					 <input type="text" name="ano_empr_devolucao" size="4" maxlength="4" value="<%=ano_devolucao%>" onKeyup="mov_devol();" id="ano_movimento"></td>
</tr>

<%elseif str_situacao = 9 then
'*** Efetivação do comodato ***%>
<tr>
	<td><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"><input type="hidden" name="cd_unidade_2" value="<%=cd_unidade2%>">
	<input type="hidden" name="cd_rack_2" value="<%=cd_rack2%>"></td>
	<td>&nbsp; Data efetivacao: <input type="text" name="dia_efetivacao" size="2" maxlength="2" value="<%'=dia_devolucao%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="mes_efetivacao" size="2" maxlength="2" value="<%=mes_efetivacao%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="ano_efetivacao" size="4" maxlength="4" value="<%=ano_efetivacao%>">
	<%="{"&cd_unidade&"-"&cd_rack&"}{"&cd_unidade2&"-"&cd_rack2&"}"%></td>
</tr>
<%end if%>
</form-->
</table>
<%end if%>
