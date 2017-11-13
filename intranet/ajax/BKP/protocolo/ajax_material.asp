<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->

<%cd_material = request("cd_material")

If cd_material = "" Then '*** Mostra campo normal ***%>
	&nbsp;
	<input type="hidden" name="cd_material_1" value="">	<!--Em branco(ajax_material.asp)-->
<%Elseif IsNumeric(cd_material) then ' *** Mostra Campo numérico ***%>
	<%xsql = "SELECT * FROM TBL_material where a061_codmat='"&cd_material&"' Order by a061_nommat" '*** Mostra pelo nome
	Set rs_mat = dbconn.execute(xsql)%>
	<%while not rs_mat.EOF%>
	<%cd_cod_mat = rs_mat("a061_codmat")
	nm_material = rs_mat("a061_nommat")%>
	
	<%'=left(rs_proc("nm_procedimento"),25)%>
	<%=cd_cod_mat&" - "&left(nm_material,75)%><%if len(nm_material) > 75 then%><%="..."%><%end if%>
	
	<%rs_mat.MoveNext
	mat_sel = ""
	wend%>
	<input type="hidden" name="cd_material_1" value="">
	 <%rs_mat.close
	set rs_mat = nothing%>

<%Elseif not IsNumeric(cd_material) then%>
		<%xsql = "SELECT * FROM TBL_material where a061_nommat like '"&cd_material&"%' order by a061_nommat"
		Set rs_mat = dbconn.execute(xsql)%>	
		<select name="cd_material_1">
		<%while not rs_mat.EOF' then%>
		<%cd_cod_mat = rs_mat("a061_codmat")
		nm_material = rs_mat("a061_nommat")%>
		<option value="<%=cd_cod_mat%>" class="inputs"><%=cd_cod_mat&" - "&left(nm_material,25)%></option>
		<%rs_mat.MoveNext
		wend' if%>
		</select>
		<input type="Button" name="mater" value="ok" onFocus="nextfield ='cd_material';" onclick="soma_mat('',document.form.cd_material_1.value,document.form.cd_material_2.value,document.form.materiais.value);">
		<%rs_mat.close
		set rs_mat = nothing%>
	
<%end if%>



