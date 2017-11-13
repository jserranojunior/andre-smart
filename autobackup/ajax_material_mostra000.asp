<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->

<%cd_material_busca = request("cd_material_busca")

If cd_material_busca = "" Then '*** Mostra campo normal ***%>
	&nbsp;
	<input type="hidden" name="cd_material_1" value="">	<!--Em branco(ajax_material.asp)-->
	<input type="button" name="somar_mat" value="+" onKeyup="nextfield ='cd_patrimonio_busca';">&nbsp;
<%Elseif IsNumeric(cd_material_busca) then ' *** Mostra Campo numérico ***%>
	<%xsql = "SELECT * FROM TBL_material where a061_codmat='"&cd_material_busca&"' Order by a061_nommat" '*** Mostra pelo nome
	Set rs_mat = dbconn.execute(xsql)%>
	<%while not rs_mat.EOF%>
	<%cd_cod_mat = rs_mat("a061_codmat")
	nm_material = rs_mat("a061_nommat")%>
	
	<%'=left(rs_proc("nm_procedimento"),25)%>
	<%=cd_cod_mat&" - "&left(nm_material,75)%><%if len(nm_material) > 75 then%><%="..."%><%end if%>
	
	<%rs_mat.MoveNext
	mat_sel = ""
	wend%>
	<input type="hiddens" name="cd_material_1" value="<%=cd_cod_mat%>">
	<input type="button" name="somar_mat" id="somar_mat" value="+"  onFocus="soma_mat(<%=cd_cod_mat%>,document.getElementById('qtd_material').value,'',document.getElementById('materiais_total').value)" >&nbsp;
	 <%rs_mat.close
	set rs_mat = nothing%>

<%Elseif not IsNumeric(cd_material_busca) then%>
		<%xsql = "SELECT * FROM TBL_material where a061_nommat like '"&busca_inteligente(cd_material_busca)&"%' order by a061_nommat"
		Set rs_mat = dbconn.execute(xsql)%>	
		<!--select name="cd_material_1" onBlur="soma_mat(document.form.cd_material_1.value,document.form.cd_material_2.value,document.form.qtd_material.value,document.form.materiais_total.value)" onKeyup="nextfield ='cd_material_busca';"-->
		<select name="cd_material_1" id="cd_material_1"  onKeyup="nextfield ='cd_material_busca';">
		<%if not rs_mat.EOF then%>			
			<%while not rs_mat.EOF' then%>
			<%cd_cod_mat = rs_mat("a061_codmat")
			nm_material = rs_mat("a061_nommat")%>
			<option value="<%=cd_cod_mat%>" class="inputs"><%=cd_cod_mat&" - "&left(nm_material,25)%></option>
			<%rs_mat.MoveNext
			wend' if%>
		<%else%>
			<option value="" class="inputs">* Material não encontrado *</option>
		<%end if%>	
		</select>
			<!--input type="Button" name="somar_mat" value="Tecle ENTER para selecionar" onFocus="nextfield ='cd_material_1';" onclick="soma_mat(document.form.cd_material_1.value,document.form.cd_material_2.value,document.form.qtd_material.value,document.form.materiais.value);"-->
			<input type="button" name="somar_mat" id="somar_mat" value="+" onFocus="soma_mat(document.getElementById('cd_material_1').value,document.getElementById('qtd_material').value,'',document.getElementById('materiais_total').value)">
			<%rs_mat.close
			set rs_mat = nothing%>
	
<%end if%>



