<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%tipo_item = request("tipo_item")
tipo = mid(tipo_item,1,1)
cod_os = mid(tipo_item,3,len(tipo_item))
response.write("ajax/ajax_patrimonio_compra.asp")

'if len(cd_patrimonio) > 2 Then
'xsql = "SELECT * FROM View_patrimonio_lista Where no_patrimonio = "&no_patrimonio&" and nm_tipo = '"&nm_tipo&"'"
'Set rs = dbconn.execute(xsql)%>
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
			<table width="400" align="left">
			<%if tipo = "e" then%>
				<tr bgcolor="#f5f5f5">
					<td  align="center">Equipamento</td>
					<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 1 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td>
						<input type="hidden" name="cd_patrimonio" value="">
						<input type="hidden" name="cd_especialidade" value="">
						<input type="hidden" name="cd_ns" value="">
						<span id='divEquip'><select name="cd_equipamento" onFocus="nextfield ='cd_marca';">
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
						  	wend%></select>
						</span>
						<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_equip();"-->
						<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=180'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
					</td>
					<td  align="center">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td><span id='divMarca'><select name="cd_marca" class="inputs" onFocus="nextfield ='cd_ns';">
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
					<!--td align="center">Especialidade</td>
								<%strsql ="Select * From TBL_especialidades where cd_codigo = 13 OR cd_codigo = 14 OR cd_codigo = 22 Order by nm_especialidade"
							  	Set rs_espec = dbconn.execute(strsql)%>
					<td><select name="cd_especialidade" onFocus="nextfield ='cd_equipamento';">
										<option class="inputs" value="">Selecione</option>
										<%Do while Not rs_espec.EOF%><%especialidade=rs_espec("nm_especialidade")%><%if nm_especialidade=especialidade Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
										<option class="inputs" value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option>	<%rs_espec.movenext
									  	Loop%>
					</td-->
					<td  align="center">Qtd	</td>
					<td><input type="text" name="num_qtd" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';"></td>
					
				</tr>
			<%elseif tipo = "o" then%>
				<tr bgcolor="#f5f5f5">
					<td  align="center">Ótica</td>
					<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 2 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td>
						<input type="hidden" name="cd_patrimonio" value="">
						<input type="hidden" name="cd_especialidade" value="">
						<input type="hidden" name="cd_ns" value="">
						<span id='divEquip'><select name="cd_equipamento" onFocus="nextfield ='cd_marca';">
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
						  	wend%></select>
						</span>
						<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_equip();"-->
						<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=180'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
					</td>
					<td  align="center">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td><span id='divMarca'><select name="cd_marca" class="inputs" onFocus="nextfield ='cd_ns';">
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
					<!--td align="center">Especialidade</td>
								<%strsql ="Select * From TBL_especialidades where cd_codigo = 13 OR cd_codigo = 14 OR cd_codigo = 22 Order by nm_especialidade"
							  	Set rs_espec = dbconn.execute(strsql)%>
					<td><select name="cd_especialidade" onFocus="nextfield ='cd_equipamento';">
										<option class="inputs" value="">Selecione</option>
										<%Do while Not rs_espec.EOF%><%especialidade=rs_espec("nm_especialidade")%><%if nm_especialidade=especialidade Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
										<option class="inputs" value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option>	<%rs_espec.movenext
									  	Loop%>
					</td-->
					<td  align="center">Qtd	</td>
					<td><input type="text" name="num_qtd" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';"></td>
					
				</tr>
			<%elseif tipo = "i" then%>
				
				<tr bgcolor="#f5f5f5">
					<td  align="center">Instrumento</td>
					<%strsql = "SELECT * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 3 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td>
						<span id='divEquip'><select name="cd_equipamento" onFocus="nextfield ='cd_marca';">
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
						  	wend%></select>
						</span>
						<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_equip();"-->
						<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=180'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
					</td>
					<td  align="center">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td><span id='divMarca'><select name="cd_marca" class="inputs" onFocus="nextfield ='cd_ns';">
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
					<td align="center">Especialidade</td>
								<%strsql ="Select * From TBL_especialidades where cd_codigo = 13 OR cd_codigo = 14 OR cd_codigo = 22 Order by nm_especialidade"
							  	Set rs_espec = dbconn.execute(strsql)%>
					<td><select name="cd_especialidade" onFocus="nextfield ='cd_equipamento';">
										<option class="inputs" value="">Selecione</option>
										<%Do while Not rs_espec.EOF%><%especialidade=rs_espec("nm_especialidade")%><%if nm_especialidade=especialidade Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
										<option class="inputs" value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option>	<%rs_espec.movenext
									  	Loop%>
					</td>
					<td  align="center">Qtd	</td>
					<td><input type="text" name="num_qtd" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';"></td>
					
				</tr>
			<%elseif tipo = "m" then%>
				
				<tr bgcolor="#f5f5f5">
					<td  align="center">Material</td>
					<%strsql = "SELECT   * FROM TBL_equipamento WHERE cd_status = 1 AND cd_tipo = 4 ORDER BY nm_equipamento_novo"
					Set rs_equip = dbconn.execute(strsql)%>
					<td>
						<span id='divEquip'><select name="cd_equipamento" onFocus="nextfield ='cd_marca';">
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
						  	wend%></select>
						</span>
						<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_equip();"-->
						<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/equipamento_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=490, height=180'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->
					</td>
					<td  align="center">Marca</td>
						<%strsql ="Select * From TBL_marca Order by nm_marca"
					  	Set rs_marca = dbconn.execute(strsql)%>
					<td><span id='divMarca'><select name="cd_marca" class="inputs" onFocus="nextfield ='cd_ns';">
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
					<td  align="center">Qtd	</td>
					<td><input type="text" name="num_qtd" size="4" maxlength="4" class="inputs" value="<%=num_qtd%>" onFocus="nextfield ='motivo';"></td>
					<td colspan="2">&nbsp;</td>
				</tr>	 
			<%end if%>
				<tr>
					<td><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="250" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
					<td><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
				</tr>
			 </table>
		</span>
	</td>
</tr>