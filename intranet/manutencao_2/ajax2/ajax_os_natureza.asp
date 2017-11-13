<!--#include file="../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%natureza_os = request("natureza_os")
natureza = mid(natureza_os,1,1)
cod_os = mid(natureza_os,3,len(natureza_os))
'response.write(natureza)
if cd_user = 46 then response.write("ajax2/ajax_os_natureza.asp")

	strsql ="SELECT * FROM TBL_OrdemServico2 WHERE cd_codigo = '"&cod_os&"'"
	Set rs = dbconn.execute(strsql)
	if not rs.EOF then
		cd_equipamento = rs("cd_equipamento")
		cd_marca = rs("cd_marca")
		cd_ns = rs("cd_ns")
		num_qtd = rs("num_qtd")
	end if%>

<%if natureza = "m" Then%>
	<table>
		<tr>
			<td style="width:110px;"><input type="radio" name="cd_tipo_item" value="e" onMouseup="seleciona_tipo_objeto_m('e');">Equipamento</td>
			<td style="width:80px;"><input type="radio" name="cd_tipo_item" value="o" onMouseup="seleciona_tipo_objeto_m('o');">Ótica</td> 
			<td style="width:110px;"><input type="radio" name="cd_tipo_item" value="i" onMouseup="seleciona_tipo_objeto_m('i');">Instrumento</td>
			<td style="width:190px;"><input type="radio" name="cd_tipo_item" value="m" onMouseup="seleciona_tipo_objeto_m('m');">Material</td>
		</tr>
	</table>
	<span id="divFase1" style="color:red;"> *** Selecione um item acima ***	</span>
<%Elseif natureza = "c" then%>
	<table>
		<tr>
			<td style="width:110px;"><input type="radio" name="cd_tipo_item" value="e" onMouseup="seleciona_tipo_objeto_c('e');">Equipamento</td>
			<td style="width:80px;"><input type="radio" name="cd_tipo_item" value="o" onMouseup="seleciona_tipo_objeto_c('o');">Ótica</td> 
			<td style="width:110px;"><input type="radio" name="cd_tipo_item" value="i" onMouseup="seleciona_tipo_objeto_c('i');">Instrumento</td>
			<td style="width:195px;"><input type="radio" name="cd_tipo_item" value="m" onMouseup="seleciona_tipo_objeto_c('m');">Material</td>
		</tr>
	</table>
	<span id="divFase1" style="color:red;"> *** Selecione um item acima ***	</span>
<%Elseif natureza = "p" then%>
	>>> Realização de Manutenção Preventiva em Equipamentos <<<
<%else%>
<%=natureza%> + <%=cod_os%> ?
<%end if%>
