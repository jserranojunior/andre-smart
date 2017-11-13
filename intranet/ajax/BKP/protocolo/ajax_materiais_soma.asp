<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

materiais = request("materiais")

materiais = "100"
'response.write(procedimentos&"<br>")

	'*** Apaga vírgula se for o último caractere ***
	if right(materiais,1) = "," Then 
	mat_len = len(materiais)
	mat_len = mat_len - 1
			materiais = Left(materiais,mat_len)
	end if



mat_array = split(materiais,",")
		For Each mat_item In mat_array
		
		if IsNumeric(mat_item) then
			xsql = "SELECT * FROM TBL_materiais where a061_codmat="&mat_item&" order by a061_nommat"
			Set rs_mat1 = dbconn.execute(xsql)
			while not rs_mat1.EOF
			
			cd_cod_mat = rs_mat1("a061_codmat")
			nm_material = rs_mat1("a061_nommat")%>
			
			<%=cd_cod_mat%> - <%=nm_material%>&nbsp; <br>
			<input type="hidden" name="materiais" value="<%=materiais%>">
			<%rs_mat1.MoveNext
			 wend
			
			  rs_mat1.close
			  set rs_mat1 = nothing
		end if
		'numero = numero + 1
		
		Next

%>