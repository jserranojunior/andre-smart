<!--#include file="../../../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

cd_cid = request("cd_cid")
		cd_sexo = mid(cd_cid,1,1)
		cd_cid = mid(cd_cid,2,len(cd_cid))
		 cd_cid = replace(cd_cid,".","")

If cd_cid = "" Then '*** Mostra campo normal ***%>
	<%xsql = "SELECT * FROM TBL_cid10_subcategoria Order by subcategoria"
	Set rs_cid = dbconn.execute(xsql)%>
	<!-- NÃO MOSTRA NADA! -->	
	
<%Elseif IsNumeric(mid(cd_cid,2,2)) then ' *** Mostra as subcategorias ***%>
	<%xsql = "SELECT top 10 * FROM TBL_cid10_subcategoria where subcategoria Like '"&cd_cid&"%' Order by subcategoria" '*** Mostra o nome da subcategoria
		Set rs_cid = dbconn.execute(xsql)%>
		<!--"select size="1" id="lista" name="cd_especialidade_1"-->
		<%if not rs_cid.EOF Then%>
			<b style="color:white; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - SUB-CATEGORIA </b><br>		
			<%while not rs_cid.EOF%>
			<%cd_cod_espec = rs_cid("cd_codigo")
			subcategoria = rs_cid("subcategoria")
			descricao = rs_cid("descricao")
			rest_sexo = rs_cid("rest_sexo")
				if int(trim(rest_sexo)) <> int(cd_sexo) then
					aviso = "<b style=color:red;> CID incompatível com o sexo </b>"
				else
					aviso = ""
				end if
				
				if int(rest_sexo) = 1 Then
					restricao = "&nbsp;(Masculino)"
				elseif int(rest_sexo) = 2 Then
					restricao = "&nbsp;(Feminino)"
				else
					restricao = ""
				end if
				
				alt_descricao = descricao
				if len(descricao) > 77 then
					descricao = mid(descricao,1,80)&"..."
				end if
			%>
			<%=mid(subcategoria,1,3)%><%if int(len(trim(subcategoria))) > 3 then response.write("."&mid(subcategoria,4,4)) end if%> - <%=descricao%> <%=rest_sexo%> <%=restricao%> <%=aviso%><br>
			<%rs_cid.MoveNext
			espec_sel = ""
			wend
		else%>
			<b style="color:white; border:1 solid gray; background-color:gray; width:587px;">CÓD. &nbsp; - &nbsp;SUB-CATEGORIA </b><br>
			* CID não encontrado *
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