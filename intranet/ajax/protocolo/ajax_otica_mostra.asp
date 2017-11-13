<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_patrimonio_busca = request("cd_patrimonio_busca")

If cd_patrimonio_busca = "" Then '*** Mostra campo normal ***%>
	&nbsp;
	<input type="text" name="cd_otica_1" value="">
	<input type="Button" name="somar" value="+" onFocus="nextfield ='cd_material_busca';">
<%Elseif IsNumeric(cd_patrimonio_busca) then ' *** Mostra Campo numérico ***
	cd_otica = left(cd_patrimonio_busca,11)%>
	<%xsql = "SELECT * FROM TBL_patrimonio where cd_codigo='"&cd_patrimonio_busca&"' AND cd_status = 1 Order by cd_codigo" '*** Mostra pelo nome
	Set rs_otic = dbconn.execute(xsql)
	if not rs_otic.EOF Then%>
		<%while not rs_otic.EOF%>
		<%cd_cod_proc = rs_otic("cd_codigo")
		nm_otica = rs_otic("nm_otica")%>
		<%'=left(rs_otic("nm_otica"),25)%>
		<%=cd_cod_proc&" - "&left(rs_otic("nm_otica"),75)%><%if len(nm_otica) > 75 then%><%="..."%><%end if%>
		<%rs_otic.MoveNext
		espec_sel = ""
		wend
	else%>
		<b style="color:red;">* Código AMB não encontrado *</b>
	<%end if%>
	<input type="hidden" name="cd_otica_1" value="<%=cd_cod_proc%>">
	<input type="button" name="somar" value="+" onFocus="soma(document.form.cd_otica_1.value,document.form.cd_otica_2.value,document.form.oticas_total.value)" onKeyup="nextfield ='cd_otica_busca';">&nbsp;

	 <%rs_otic.close
	set rs_otic = nothing%>

<%Elseif not IsNumeric(cd_patrimonio_busca) then%>
		<%xsql = "SELECT * FROM TBL_otica where nm_otica like '"&cd_patrimonio_busca&"%' AND cd_status = 1 order by nm_otica"
		Set rs_otic = dbconn.execute(xsql)%>	
			<select name="cd_otica_1" onBlur="soma(document.form.cd_otica_1.value,document.form.cd_otica_2.value,document.form.oticas_total.value)" onKeyup="nextfield ='cd_otica_busca';">
		<%if not rs_otic.EOF Then%>
			<%while not rs_otic.EOF' then%>
			<%cd_cod_proc = rs_otic("cd_codigo")
			nm_otica = rs_otic("nm_otica")%>
			<option value="<%=cd_cod_proc%>" class="inputs"><%=cd_cod_proc&" - "&left(rs_otic("nm_otica"),25)%></option>
			<%rs_otic.MoveNext
			wend' if%>
		<%else%>
		<option value="" class="inputs">* otica não encontrado *</option>
		<%end if%>
	</select>
		<input type="Button" name="somar" value="Tecle ENTER para selecionar" onFocus="nextfield ='cd_otica_1';">
		<%rs_otic.close
		set rs_otic = nothing%>
	
<%end if%>



