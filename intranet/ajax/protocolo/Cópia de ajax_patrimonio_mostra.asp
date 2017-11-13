<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="pt-br" lang="pt-br">
<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->





<%'txtBusca = request("txtBusca")

cd_patrimonio_busca = request("cd_patrimonio_busca")

If cd_patrimonio_busca = "" Then '*** Mostra campo normal ***%>
	<!--&nbsp;	
	<input type="Button" name="somar" value="+" onFocus="nextfield ='cd_material_busca';"-->
	
	<%'cd_otica = left(cd_patrimonio_busca,11)%>
	<%'xsql = "SELECT * FROM View_patrimonio_lista where cd_patrimonio='"&cd_patrimonio_busca&"' AND cd_status = 1 Order by cd_patrimonio" '*** Mostra pelo nome
	'xsql = "SELECT * FROM View_patrimonio_lista where cd_status = 1 Order by cd_patrimonio" '*** Mostra pelo nome
	'Set rs_patr = dbconn.execute(xsql)
	'if not rs_patr.EOF Then%>
		<%'while not rs_patr.EOF%>
		<%'cd_patrimonio = rs_patr("cd_patrimonio")
		'nm_equipamento = rs_patr("nm_equipamento")%>
		<%'=left(rs_patr("nm_equipamento"),25)%>
		Vd Lap <%'=cd_patrimonio&" ( "&left(rs_patr("nm_equipamento"),75)%><%'if len(nm_equipamento) > 75 then%><%'="..."%><%'end if%>)
		<%'rs_patr.MoveNext
		'patr_sel = ""
		'wend
	'else%>
		<!--b style="color:red;">* Patrimônio não encontrado *</b-->
	<%'end if%>
	<!--input type="text" name="cd_patrimonio" value="<%=cd_patrimonio%>">	
	<input type="button" name="somar" value="+" onFocus="soma_patrimonio(document.Form.cd_patrimonio.value,'',document.Form.subtotal_patrimonio.value)" onKeyup="nextfield ='cd_otica_busca';">&nbsp;
	-->
	 <%'rs_patr.close
	'set rs_patr = nothing%>
	
	
	
<%Elseif IsNumeric(cd_patrimonio_busca) then ' *** Mostra Campo numérico ***
	cd_otica = left(cd_patrimonio_busca,11)%>
	<%xsql = "SELECT * FROM View_patrimonio_lista where cd_patrimonio='"&cd_patrimonio_busca&"' AND cd_status = 1 Order by cd_patrimonio" '*** Mostra pelo nome
	Set rs_patr = dbconn.execute(xsql)
	if not rs_patr.EOF Then%>
		<%while not rs_patr.EOF%>
		<%cd_patrimonio = rs_patr("cd_patrimonio")
		nm_equipamento = rs_patr("nm_equipamento")%>
		<%'=left(rs_patr("nm_equipamento"),25)%>
		Vd Lap (<%=cd_patrimonio&" "&left(rs_patr("nm_equipamento"),75)%><%if len(nm_equipamento) > 75 then%><%="..."%><%end if%>)
		<%rs_patr.MoveNext
		patr_sel = ""
		wend
	else%>
		<b style="color:red;">* Patrimônio não encontrado *</b>
	<%end if%>
	<input type="text" name="cd_patrimonio" value="<%=cd_patrimonio%>">	
	<input type="button" name="somar" value="+" onClick="soma_patrimonio(document.Form.cd_patrimonio.value,'',document.Form.subtotal_patrimonio.value)" onKeyup="nextfield ='cd_otica_busca';">&nbsp;

	 <%rs_patr.close
	set rs_patr = nothing%>

<%Elseif not IsNumeric(cd_patrimonio_busca) then%>
		<%xsql = "SELECT * FROM View_patrimonio_lista where nm_equipamento like '%"&cd_patrimonio_busca&"%' AND cd_status = 1 order by nm_equipamento"
		Set rs_patr = dbconn.execute(xsql)%>	
			<!--select name="cd_patrimonio" onBlur="soma_patrimonio(document.Form.cd_patrimonio.value,'',document.Form.subtotal_patrimonio.value)" onKeyup="nextfield ='cd_otica_busca';"-->
			<select name="cd_patrimonio" onKeyup="nextfield ='cd_otica_busca';">
		<%if not rs_patr.EOF Then%>
			<%while not rs_patr.EOF' then%>
			<%cd_patrimonio = rs_patr("cd_patrimonio")
			nm_equipamento = rs_patr("nm_equipamento")%>
			<option value="<%=cd_patrimonio%>" class="inputs"><%=cd_patrimonio&" - "&left(rs_patr("nm_equipamento"),25)%></option>
			<%rs_patr.MoveNext
			wend' if%>
		<%else%>
		<option value="" class="inputs">* Patrimônio não encontrado *</option>
		<%end if%>
	</select>
		<input type="Button" name="somar" value="Tecle ENTER para selecionar" onClick="soma_patrimonio(document.Form.cd_patrimonio.value,'',document.Form.subtotal_patrimonio.value)">
		<%rs_patr.close
		set rs_patr = nothing%>
	
<%end if%>



