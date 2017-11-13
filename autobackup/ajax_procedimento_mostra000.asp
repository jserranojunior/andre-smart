<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_procedimento_busca = request("cd_procedimento_busca")
'cd_procedimento_busca = replace(cd_procedimento_busca,"%20"," ")

If cd_procedimento_busca = "" Then '*** Mostra campo normal ***%>
	&nbsp;
	<input type="hidden" name="cd_procedimento_1" value="">
	<input type="Button" name="somar" value="+" onFocus="nextfield ='cd_material_busca';">
<%Elseif IsNumeric(cd_procedimento_busca) then ' *** Mostra Campo numérico ***
	cd_procedimento = left(cd_procedimento_busca,11)%>
	<%xsql = "SELECT * FROM TBL_procedimento where cd_codigo='"&cd_procedimento_busca&"' AND cd_status = 1 Order by cd_codigo" '*** Mostra pelo nome
	Set rs_proc = dbconn.execute(xsql)
	if not rs_proc.EOF Then%>
		<%while not rs_proc.EOF%>
		<%cd_cod_proc = rs_proc("cd_codigo")
		nm_procedimento = rs_proc("nm_procedimento")%>
		<%'=left(rs_proc("nm_procedimento"),25)%>
		<%=cd_cod_proc&" - "&left(rs_proc("nm_procedimento"),75)%><%if len(nm_procedimento) > 75 then%><%="..."%><%end if%>
		<%rs_proc.MoveNext
		espec_sel = ""
		wend
	else%>
		<b style="color:red;">* Código AMB não encontrado *</b>
	<%end if%>
	<input type="hidden" name="cd_procedimento_1" id="cd_procedimento_1" value="<%=cd_cod_proc%>">
	<!--input type="button" name="somar" value="+" onFocus="soma(document.form.cd_procedimento_1.value,document.form.cd_procedimento_2.value,document.form.procedimentos_total.value)" onKeyup="nextfield ='cd_procedimento_busca';"-->
	<input type="button" name="somar" id="somar_proc" value="+" onFocus="soma(<%=cd_cod_proc%>,'',document.getElementById('procedimentos_total').value)">&nbsp;

	 <%rs_proc.close
	set rs_proc = nothing%>

<%Elseif not IsNumeric(cd_procedimento_busca) then%>
		<%xsql = "SELECT * FROM TBL_procedimento where nm_procedimento like '"&busca_inteligente(cd_procedimento_busca)&"%' AND cd_status = 1 order by nm_procedimento"
		Set rs_proc = dbconn.execute(xsql)%>	
			<select name="cd_procedimento_1" onBlur="soma(this.value,'',document.getElementById('procedimentos_total').value)">
		<%if not rs_proc.EOF Then%>
			<%while not rs_proc.EOF' then%>
			<%cd_cod_proc = rs_proc("cd_codigo")
			nm_procedimento = rs_proc("nm_procedimento")%>
			<option value="<%=cd_cod_proc%>" class="inputs" onFocus><%=cd_cod_proc&" - "&left(rs_proc("nm_procedimento"),50)%></option>
			<%rs_proc.MoveNext
			wend' if%>
		<%else%>
		<option value="" class="inputs">* Procedimento não encontrado *</option>
		<%end if%>
	</select>-
		<input type="Button" name="somar" value="ENTER p/ selecionar" onFocus="nextfield ='cd_procedimento_1';">
		<%rs_proc.close
		set rs_proc = nothing%>
	
<%end if%>



