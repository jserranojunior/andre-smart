<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%tipo_item = request("tipo_item")
tipo = mid(tipo_item,1,1)
cod_os = mid(tipo_item,3,len(tipo_item))
'if cd_user = 46 then response.write("ajax/ajax_os_patrimonio_compra.asp")

'if len(cd_patrimonio) > 2 Then
'xsql = "SELECT * FROM View_patrimonio_lista Where no_patrimonio = "&no_patrimonio&" and nm_tipo = '"&nm_tipo&"'"
'Set rs = dbconn.execute(xsql)

wdt_nome_item = 75
wdt_campo_item = 333%>
<!--#include file="../js/manutencao2.js"-->
<%if cod_os <> "" Then%>
<tr><td>&nbsp;<br></td></tr>
<tr>
	<td align="center"><b>Patrimônio:</b></td>
	<td><input type="text" name="num_patrimonio" maxlength="10" size="15" value="<%=cd_patrimonio%>" class="inputs" onFocus="nextfield ='num_qtd';"  onKeyup="mostra_patrimonio('<%=tipo%>');">
</tr>
<%end if%>
<tr>
	<td colspan="4">
		<span id="divPat_mostra">
			<table>
			<%if tipo = "e" then%>
				<tr bgcolor="#f5f5f5">
					<td align="center" style="width:<%=wdt_nome_item%>px;">Equipamento</td>
					<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 1 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;">
						<span id='divEquip'>
							<select name="cd_equipamento_x" class="inputs" onFocus="nextfield ='cd_marca';" onChange="alimenta_cd_equipamento(this.value);">
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
					<td  align="center" style="width:<%=wdt_nome_item%>px;">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;"><span id='divMarca'>
							<select name="cd_marca_x" class="inputs" onFocus="nextfield ='cd_ns';" onChange="alimenta_cd_marca(this.value);">
								<option value="">Selecione</option>
								<%Do while Not rs_marca.EOF %><%marca=rs_marca("nm_marca")%><%if nm_marca=marca Then%><%ck_marca="selected"%><%else ck_marca=""%><%end if%>
								<option value="<%=rs_marca("cd_codigo")%>" <%=ck_marca%>><%=rs_marca("nm_marca")%></option>
								<%rs_marca.movenext
							  	Loop%>
							</select>
						</span>
					</td>	
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center" style="width:<%=wdt_nome_item%>px;">Qtd	</td>
					<td style="width:<%=wdt_campo_item%>px;"><input type="text" name="num_qtd_x" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';" onkeyup="alimenta_cd_qtd(this.value);"></td>
				</tr>
			<%elseif tipo = "o" then%>
				<tr bgcolor="#f5f5f5">
					<td align="center" style="width:<%=wdt_nome_item%>px;">Ótica</td>
					<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 2 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;">
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
					<td style="width:<%=wdt_campo_item%>px;"><span id='divMarca'><select name="cd_marca_x" class="inputs" onFocus="nextfield ='cd_ns';" onChange="alimenta_cd_marca(this.value);">
						<option value="">Selecione</option>
						<%Do while Not rs_marca.EOF %><%marca=rs_marca("nm_marca")%><%if nm_marca=marca Then%><%ck_marca="selected"%><%else ck_marca=""%><%end if%>
						<option value="<%=rs_marca("cd_codigo")%>" <%=ck_marca%>><%=rs_marca("nm_marca")%></option>
						<%rs_marca.movenext
					  	Loop%>
						</select>
						</span>
					</td>	
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center" style="width:<%=wdt_nome_item%>px;">Qtd	</td>
					<td style="width:<%=wdt_campo_item%>px;"><input type="text" name="num_qtd_x" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';"  onkeyup="alimenta_cd_qtd(this.value);"></td>
				</tr>
			<%elseif tipo = "i" then%>
				
				<tr bgcolor="#f5f5f5">
					<td align="center"  style="width:<%=wdt_nome_item%>px;">Instrumento</td>
					<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 3 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;">
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
					<td align="center"  style="width:<%=wdt_nome_item%>px;">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;"><span id='divMarca'>
							<select name="cd_marca_x" class="inputs" onFocus="nextfield ='cd_ns';" onChange="alimenta_cd_marca(this.value);">
								<option value="">Selecione</option>
								<%Do while Not rs_marca.EOF %><%marca=rs_marca("nm_marca")%><%if nm_marca=marca Then%><%ck_marca="selected"%><%else ck_marca=""%><%end if%>
								<option value="<%=rs_marca("cd_codigo")%>" <%=ck_marca%>><%=rs_marca("nm_marca")%></option>
								<%rs_marca.movenext
							  	Loop%>
							</select>
						</span>
					</td>	
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center"  style="width:<%=wdt_nome_item%>px;">Especialidade</td>
								<%strsql ="Select * From TBL_especialidades where cd_codigo = 13 OR cd_codigo = 14 OR cd_codigo = 22 Order by nm_especialidade"
							  	Set rs_espec = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;">
						<select name="cd_especialidade_x" class="inputs" onFocus="nextfield ='cd_equipamento';" onChange="alimenta_cd_especialidade(this.value);">
							<option class="inputs" value="">Selecione</option>
							<%Do while Not rs_espec.EOF%><%especialidade=rs_espec("nm_especialidade")%><%if nm_especialidade=especialidade Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
							<option class="inputs" value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option>
							<%rs_espec.movenext
							Loop%>
						</select>
					</td>
				</tr>
				<tr>
					<td align="center"  style="width:<%=wdt_nome_item%>px;">Qtd	</td>
					<td style="width:<%=wdt_campo_item%>px;"><input type="text" name="num_qtd_x" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';" onkeyup="alimenta_cd_qtd(this.value);"></td>
					
				</tr>
			<%elseif tipo = "m" then%>
				
				<tr bgcolor="#f5f5f5">
					<td align="center" style="width:<%=wdt_nome_item%>px;">Material</td>
					<%strsql = "SELECT   * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 4 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;">
						<span id='divEquip'>
							<select name="cd_equipamento_x" onFocus="nextfield ='cd_marca';" class="inputs" onChange="alimenta_cd_equipamento(this.value);">
								<option value="" class="inputs">Selecione</option>
								<%while Not rs_equip.EOF%>
								<%cd_equip = rs_equip("cd_codigo")
								equip = rs_equip("nm_equipamento_novo")%>
								<%if int(cd_equipamento) = cd_equip then%>
									<%ck_equip = "selected"%>
								<%else ck_equip = ""%>
								<%end if%>
								<option class="inputs" value="<%=rs_equip("cd_codigo")%>" <%=ck_equip%>><%=equip%></option>
								<%rs_equip.movenext
							  	wend%>
							</select>
						</span>
						<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_equip();"-->
						<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=180'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
					</td>
				</tr>
				<tr>
					<td  align="center" style="width:<%=wdt_nome_item%>px;">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td style="width:<%=wdt_campo_item%>px;">
						<span id='divMarca'>
							<select name="cd_marca_x" class="inputs" onFocus="nextfield ='cd_ns';" onChange="alimenta_cd_marca(this.value);">
								<option value="">Selecione</option>
								<%Do while Not rs_marca.EOF %><%marca=rs_marca("nm_marca")%><%if nm_marca=marca Then%><%ck_marca="selected"%><%else ck_marca=""%><%end if%>
								<option value="<%=rs_marca("cd_codigo")%>" <%=ck_marca%>><%=rs_marca("nm_marca")%></option>
								<%rs_marca.movenext
							  	Loop%>
							</select>
						</span>
					</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center" style="width:<%=wdt_nome_item%>px;">Qtd	</td>
					<td style="width:<%=wdt_campo_item%>px;"><input type="text" name="num_qtd_x" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';" onkeyup="alimenta_cd_qtd(this.value);"></td>
				</tr>	 
			<%end if%>
				
			 </table>
		</span>
	</td>
</tr>