<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_rack = request("cd_rack_hide")
cd_unidade = mid(cd_rack,1,instr(cd_rack,";")-1)
cd_rack = mid(cd_rack,instr(cd_rack,";")+1,len(cd_rack))

If cd_rack = "" Then '*** Mostra campo normal ***%>
	<%if cd_unidade = 20 Then
		xsql ="Select * from TBL_Rack where a056_status = 1 and a056_codrac >= 12 order by A056_desrac"' Força mostrar apenas os racks do Hosp. Edmundo
	Else
		xsql ="Select * from TBL_Rack where a056_status = 1 order by A056_desrac"' order by nm_rack"
	end if
	'xsql = "SELECT * FROM TBL_rack where a056_status = 1 Order by a056_desrac"
	Set rs_rack = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_rack_1" onChange="alimenta_rack(this.value);" onBlur="alimenta_rack(this.value);">
		<option value="" class="inputs"></option>
		<%while not rs_rack.EOF
			cd_cod_rack = rs_rack("a056_codrac")
			nm_rack = rs_rack("a056_desrac")%>	
		<option value="<%=cd_cod_rack%>" class="inputs"><%=nm_rack%></option>
		<%rs_rack.MoveNext
		wend%>
	</select>
	<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_rack(0);">
	
<%'Elseif IsNumeric(cd_rack) then ' *** Mostra Campo numérico ***
Elseif len(cd_rack) <= 2 then%>
	<%cd_rack = int(cd_rack)
	
	if cd_unidade = "20" Then
		xsql ="Select * from TBL_Rack where no_rack like '%"&busca_inteligente(cd_rack)&"%'AND  a056_status = 1 AND a056_status = 1 and a056_codrac between 11 and 13  order by A056_desrac"' 
	Else
		xsql ="Select * from TBL_Rack where no_rack like '%"&busca_inteligente(cd_rack)&"%'AND  a056_status = 1 order by A056_desrac"' order by nm_rack"
	end if
	'xsql = "SELECT * FROM TBL_rack where a056_status = 1 Order by a056_codrac" '*** Mostra pelo nome
	Set rs_rack = dbconn.execute(xsql)%>
	<select size="1" id="lista" name="cd_rack_1" onChange="alimenta_rack(this.value);" onBlur="alimenta_rack(this.value);">
		
		<%'while not rs_rack.EOF
		if not rs_rack.EOF then%>
		<%cd_cod_rack = rs_rack("a056_codrac")
		nm_rack = rs_rack("a056_desrac")
		no_rack = rs_rack("no_rack")
		'if int(cd_cod_rack) = int(cd_rack) then%>
		<option value="<%=cd_cod_rack%>" class="inputs" <%if no_rack = cd_rack then%>"selected"<%end if%> ><%=nm_rack%>&nbsp;</option>
		<%rs_rack.MoveNext
		rack_sel = ""
		'wend
		end if%>
	 </select>
<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_rack(<%=cd_cod_rack%>);">
<%Elseif not IsNumeric(cd_rack) then%>
	<%if cd_unidade = 20 Then
		'xsql ="Select * from TBL_Rack where a056_desrac like '%"&busca_inteligente(cd_rack)&"%' AND a056_status = 1  and a056_codrac >= 12 order by a056_desrac"' Força mostrar apenas os racks do Hosp. Edmundo
		xsql ="Select * from TBL_Rack where a056_desrac like '%"&busca_inteligente(cd_rack)&"%' AND a056_status = 1  and a056_codrac >= 12 order by a056_desrac"' Força mostrar apenas os racks do Hosp. Edmundo
	Else
		'xsql ="Select * from TBL_Rack where a056_desrac like '%"&busca_inteligente(cd_rack)&"%' AND a056_status = 1 order by a056_desrac"' order by nm_rack"
		xsql ="Select * from TBL_Rack where a056_desrac like '%"&busca_inteligente(cd_rack)&"%' AND a056_status = 1 order by a056_desrac"' order by nm_rack"
	end if
	'xsql = "SELECT * FROM TBL_rack where a056_desrac like '"&cd_rack&"%' AND a056_status = 1 order by a056_desrac"
	Set rs_rack = dbconn.execute(xsql)%>	
	<select size="1" id="lista" name="cd_rack_1" onChange="alimenta_rack(this.value);" onBlur="alimenta_rack(this.value);">
	<%if not rs_rack.EOF Then
		'while not rs_rack.EOF
		if not rs_rack.EOF then%>
		<%cd_cod_rack = rs_rack("a056_codrac")
		nm_rack = rs_rack("a056_desrac")%>
		<option value="<%=cd_cod_rack%>" class="inputs"><%=nm_rack%></option>
		<%rs_rack.MoveNext
		end if
		'wend%>
	<%else%>
	<option value="" class="inputs" style="color:red;">* Rack não encontrado *</option>
	<%end if%>
	</select>
	<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="alimenta_rack(<%=cd_cod_rack%>);">
<%end if%>


<%rs_rack.close
set rs_rack = nothing%>