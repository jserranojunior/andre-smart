<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%tipo_item = request("tipo_item")
'if cd_user = 46 then response.write("ajax/ajax_os_patrimonio_busca.asp")
'response.write(tipo_item)
response.write("ajax2/ajax_os_patrimonio_busca.asp")

wdt_nome_item = 150
wdt_campo_item = 200%>
<!--#include file="../js/manutencao2.js"-->
<table width="400" align="left">
<%if tipo_item = "e" then%>
	<table>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;"><b>Patrimônio</b></td>
			<td align="left" style="width:<%=wdt_nome_item%>px;"><input type="text" name="num_patrimonio" maxlength="30" size="15" value="" id="numPAT" class="inputs" onKeyup="mostra_pat('e',this.value);"></td>
		</tr>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;"><b>Numero Série:</b></td>
			<td align="left" style="width:<%=wdt_nome_item%>px;"><input type="text" name="num_serie" maxlength="30" size="15" value="" id="numNS" class="inputs" onKeyup="mostra_NS('e',this.value);"></td>
		</tr>
	</table>	
		<span id="divPatNS_mostra">&nbsp;</span>
		
<%elseif tipo_item = "o" then%>
	<table>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;"><b>Patrimônio</b></td>
			<td align="left" style="width:<%=wdt_nome_item%>px;"><input type="text" name="num_patrimonio" maxlength="30" size="15" value="" id="numPAT" class="inputs" onKeyup="mostra_pat('o',this.value);"></td>
		</tr>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;"><b>Numero Série:</b></td>
			<td align="left" style="width:<%=wdt_nome_item%>px;"><input type="text" name="num_serie" maxlength="30" size="15" value="" id="numNS" class="inputs" onKeyup="mostra_NS('o',this.value);"></td>
		</tr>
	</table>	
		<span id="divPatNS_mostra">&nbsp;</span>
		
<%'************************************
'*** Secção I N S T R U M E N T O S ***
'**************************************
elseif tipo_item = "i" then%>
	<table>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;">Instrumento</td>
			<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 3 ORDER BY nm_equipamento_novo"
			Set rs_equip = dbconn.execute(strsql)%>
			<td>
				<!--span id='divEquip'-->
					<select name="cd_equipamento_x" class="inputs" onFocus="nextfield ='cd_marca';" onChange="alimenta_cd_equipamento(this.value);">
						<option value="" class="inputs">Selecione</option>
						<%while Not rs_equip.EOF%>
						<%cd_equip = rs_equip("cd_codigo")
						equip = rs_equip("nm_equipamento_novo")%>
						<option value="<%=rs_equip("cd_codigo")%>" <%if int(cd_equipamento) = cd_equip then response.write("selected")%> ><%=equip%></option>
						<%rs_equip.movenext
					  	wend%>
					</select>
			</td>
		</tr>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;">Marca</td>
			<%strsql ="Select * From TBL_marca Order by nm_marca"
		  	Set rs_marca = dbconn.execute(strsql)%>
			<td>
				<span id='divMarca'>
					<select name="cd_marca_x" class="inputs" onFocus="nextfield ='cd_ns';" onChange="alimenta_cd_marca(this.value);">
						<option value="">Selecione</option>
						<%while Not rs_marca.EOF 
						marca = rs_marca("nm_marca")%>
						<option value="<%=rs_marca("cd_codigo")%>" <%if nm_marca=marca Then response.write("selected")%>><%=rs_marca("nm_marca")%></option>
						<%rs_marca.movenext
					  	wend%>
					</select>
				</span>
			</td>	
		</tr>
		<tr>
			<td align="center" style="width:<%=wdt_nome_item%>px;">Especialidade</td>
			<%strsql ="Select * From TBL_especialidades where cd_codigo = 13 OR cd_codigo = 14 OR cd_codigo = 22 Order by nm_especialidade"
			Set rs_espec = dbconn.execute(strsql)%>
			<td>
				<select name="cd_especialidade_x" class="inputs" onFocus="nextfield ='cd_equipamento';" onChange="alimenta_cd_especialidade(this.value);">
					<option value="">Selecione</option>
					<%while Not rs_espec.EOF
					especialidade = rs_espec("nm_especialidade")%>
					<option class="inputs" value="<%=rs_espec("cd_codigo")%>" <%if nm_especialidade=especialidade Then response.write("selected")%>><%=rs_espec("nm_especialidade")%></option>	
					<%rs_espec.movenext
					wend%>
				</select>
			</td>
		</tr>
		<tr>
			<td  align="center"style="width:<%=wdt_nome_item%>px;">Qtd	</td>
			<td style="width:<%=wdt_nome_item%>px;">
				<input type="text" name="num_qtd_x" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';" onkeyup="alimenta_cd_qtd(this.value);">
			</td>
		</tr>
	</table>
<%elseif tipo_item = "m" then%>
	
	<tr bgcolor="#f5f5f5">
		<td align="center" style="width:<%=wdt_nome_item%>px;">Material</td>
		<%strsql = "SELECT   * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 4 ORDER BY nm_equipamento_novo"
		Set rs_equip = dbconn.execute(strsql)%>
		<td>
			<span id='divEquip'>
				<select name="cd_equipamento_x" onFocus="nextfield ='cd_marca';" class="inputs" onChange="alimenta_cd_equipamento(this.value);">
					<option value="" class="inputs">Selecione</option>
					<%while Not rs_equip.EOF%>
					<%cd_equip = rs_equip("cd_codigo")
					equip = rs_equip("nm_equipamento_novo")%>
					<%'if nm_equipamento=equip Then
					if int(cd_equipamento) = cd_equip then%>
						<%ck_equip = "selected"%>
					<%else ck_equip = ""%>
					<%end if%>
					<option class="inputs" value="<%=rs_equip("cd_codigo")%>" <%=ck_equip%>><%=equip%></option>
					<%rs_equip.movenext
				  	wend%>
				</select>
			</span>
		</td>
	</tr>
	<tr>
		<td align="center" style="width:<%=wdt_nome_item%>px;">Marca</td>
			<%strsql ="Select * From TBL_marca Order by nm_marca"
		  	Set rs_marca = dbconn.execute(strsql)%>
		<td style="width:<%=wdt_nome_item%>px;"><span id='divMarca'><select name="cd_marca_x" class="inputs" onFocus="nextfield ='cd_ns';" onChange="alimenta_cd_marca(this.value);">
		<option value="">Selecione</option>
			<%Do while Not rs_marca.EOF %><%marca=rs_marca("nm_marca")%><%if nm_marca=marca Then%><%ck_marca="selected"%><%else ck_marca=""%><%end if%>
			<option value="<%=rs_marca("cd_codigo")%>" <%=ck_marca%>><%=rs_marca("nm_marca")%></option>
			<%rs_marca.movenext
		  	Loop%>
			</select>
			</span>
			<!--img src="../imagens/reload6.gif" type="button" alt="Atualizar" border="0" onClick="fill_marca();"--> 
			<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/marca_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=160'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
		</td>
	</tr>
	<tr bgcolor="#f5f5f5">
		<td align="center" style="width:<%=wdt_nome_item%>px;">Qtd	</td>
		<td style="width:<%=wdt_nome_item%>px;"><input type="text" name="num_qtd" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';" onkeyup="alimenta_cd_qtd(this.value);"></td>
		<td colspan="2">&nbsp;</td>
	</tr>	 
<%end if%>
	<!--tr>
		<td><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="250" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
		<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
	</tr-->
 </table>
