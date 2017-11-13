<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_procedimento = request("cd_procedimento")

If cd_procedimento = "" Then '*** Mostra campo normal ***%>
	&nbsp;
	<input type="hidden" name="cd_procedimento_1" value="">	
<%Elseif IsNumeric(cd_procedimento) then ' *** Mostra Campo numérico ***
	cd_procedimento = left(cd_procedimento,11)%>
	<%xsql = "SELECT * FROM TBL_procedimento where cd_codigo='"&cd_procedimento&"' Order by cd_codigo" '*** Mostra pelo nome
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
		* Código AMB não encontrado *
	<%end if%>
	<input type="hidden" name="cd_procedimento_1" value="">
	 <%rs_proc.close
	set rs_proc = nothing%>

<%Elseif not IsNumeric(cd_procedimento) then%>
		<%xsql = "SELECT * FROM TBL_procedimento where nm_procedimento like '"&cd_procedimento&"%' order by nm_procedimento"
		Set rs_proc = dbconn.execute(xsql)%>	
			<select name="cd_procedimento_1">
		<%if not rs_proc.EOF Then%>
			<%while not rs_proc.EOF' then%>
			<%cd_cod_proc = rs_proc("cd_codigo")
			nm_procedimento = rs_proc("nm_procedimento")%>
			<option value="<%=cd_cod_proc%>" class="inputs"><%=cd_cod_proc&" - "&left(rs_proc("nm_procedimento"),25)%></option>
			<%rs_proc.MoveNext
			wend' if%>
		<%else%>
		<option value="" class="inputs">* Procedimento não encontrado *</option>
		<%end if%>
	</select>
		<input type="Button" name="proced" value="ok" onFocus="nextfield ='cd_procedimento';" onclick="soma('',document.form.cd_procedimento_1.value,document.form.cd_procedimento_2.value,document.form.procedimentos.value);">
		<%rs_proc.close
		set rs_proc = nothing%>
	
<%end if%>



