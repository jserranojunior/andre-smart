<!--#include file="../../../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_cid = request("cd_cid")


If cd_cid = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_cid10_subcategoria Order by subcategoria"
	Set rs_cid = dbconn.execute(xsql)%>
	<!--select size="1" id="lista" name="cd_especialidade_1">
	<option value="" class="inputs"></option>
	<%'while not rs_cid.EOF%>
	<%'cd_cod_cid = rs_cid("cd_codigo")
	'subcategoria = rs_cid("subcategoria")
	'descricao = rs_cid("descricao")
	%>	
	<option value="<%'=cd_cod_cid%>" class="inputs"><%'=descricao%>*</option>
	<%'rs_cid.MoveNext
	'wend%>
	</select-->
	
<%Elseif IsNumeric(mid(cd_cid,2,2)) then ' *** Mostra as subcategorias ***%>
	<%if len(cd_cid) > 0 then
		xsql = "SELECT top 10 * FROM TBL_cid10_subcategoria where subcategoria Like '"&cd_cid&"%' Order by subcategoria" '*** Mostra o nome da subcategoria
		Set rs_cid = dbconn.execute(xsql)%>
		<!--"select size="1" id="lista" name="cd_especialidade_1"-->
		<%if not rs_cid.EOF Then%>
			<b style="color:white; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - SUB-CATEGORIA </b><br>		
			<%while not rs_cid.EOF%>
			<%cd_cod_espec = rs_cid("cd_codigo")
			subcategoria = rs_cid("subcategoria")
			descricao = rs_cid("descricao")
				alt_descricao = descricao
				if len(descricao) > 77 then
					descricao = mid(descricao,1,80)&"..."
				end if
			%>
			<%=mid(subcategoria,1,3)%><%if int(len(subcategoria)) > 3 then response.write("."&mid(subcategoria,4,4))%> - <%=descricao%><br>
			<%rs_cid.MoveNext
			espec_sel = ""
			wend
		else%>
			<b style="color:white; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - &nbsp;SUB-CATEGORIA </b><br>
			* CID não encontrado *
		<%end if
	else '*** Mostra as categorias ***
		xsql = "SELECT top 10 * FROM TBL_cid10_categoria where categoria Like '"&cd_cid&"%' Order by categoria" '*** Mostra o nome da categoria
		Set rs_cid = dbconn.execute(xsql)%>
		<!--"select size="1" id="lista" name="cd_especialidade_1"-->
		<%if not rs_cid.EOF Then%>
			<b style="color:yellow; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - &nbsp;CATEGORIA </b><br>
			<%while not rs_cid.EOF%>
			<%cd_cod_espec = rs_cid("cd_codigo")
			categoria = rs_cid("categoria")
			descricao = rs_cid("descricao")
				alt_descricao = descricao
				if len(descricao) > 77 then
					descricao = mid(descricao,1,80)&"..."
				end if
			%>
			<%=mid(categoria,1,3)%> - <%=descricao%><br>
			<%rs_cid.MoveNext
			espec_sel = ""
			wend
		else%>
			<b style="color:yellow; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - &nbsp;CATEGORIA </b><br>
			* CID não encontrado *
		<%end if%>		
	<%end if%>

<%Elseif not IsNumeric(mid(cd_cid,2,2)) then%>
		<%xsql = "SELECT * FROM TBL_cid10_subcategoria where descricao like '"&cd_cid&"%' order by descricao"
		Set rs_cid = dbconn.execute(xsql)%>		
	<%if not rs_cid.EOF Then%>
		<%while not rs_cid.EOF%>
		<%cd_cod_espec = rs_cid("cd_codigo")
	subcategoria = rs_cid("subcategoria")
	descricao = rs_cid("descricao")
		alt_descricao = descricao
		if len(descricao) > 77 then
			descricao = mid(descricao,1,80)&"..."
		end if%>
			<%=mid(subcategoria,1,3)%>.<%=mid(subcategoria,4,4)%> - <%=descricao%><br>
		<%rs_cid.MoveNext
		wend%>
	<%else%>
		<b style="color:yellow; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - &nbsp;CATEGORIA </b><br>
		* CID não encontrado *
	<%end if%>
		
	
	
<%end if%>


<%rs_cid.close
set rs_cid = nothing%>