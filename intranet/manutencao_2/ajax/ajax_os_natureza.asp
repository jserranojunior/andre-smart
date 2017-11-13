<!--#include file="../../includes/inc_open_connection.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%natureza_os = request("natureza_os")
natureza = mid(natureza_os,1,1)
cod_os = mid(natureza_os,3,len(natureza_os))

	strsql ="SELECT * FROM TBL_OrdemServico2 WHERE cd_codigo = '"&cod_os&"'"
	Set rs = dbconn.execute(strsql)
	if not rs.EOF then
		cd_equipamento = rs("cd_equipamento")
		cd_marca = rs("cd_marca")
		cd_ns = rs("cd_ns")
		num_qtd = rs("num_qtd")
	end if

if natureza = "m" Then%>
	<input type="radio" name="cd_tipo_item" value="e" onFocus="mostra_itens_m('e');">Equipamento &nbsp; &nbsp; &nbsp; &nbsp; 
	<input type="radio" name="cd_tipo_item" value="o" onFocus="mostra_itens_m('o');">Ótica &nbsp; &nbsp; &nbsp; &nbsp; 
	<input type="radio" name="cd_tipo_item" value="i" onFocus="mostra_itens_m('i');">Instrumento &nbsp; &nbsp; &nbsp; &nbsp; 
	<input type="radio" name="cd_tipo_item" value="m" onFocus="mostra_itens_m('m');">Material
		<br>&nbsp;
		<span id="divFase1">
		
			<input type="hidden" name="cd_patrimonio" value="<%=cd_patrimonio%>">
			<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
			<input type="hidden" name="cd_marca" value="<%=cd_marca%>">
			<input type="hidden" name="cd_especialidade" value="<%=cd_especialidade%>">
			<input type="hidden" name="cd_ns" value="<%=cd_ns%>">
			<input type="hidden" name="num_qtd" value="<%=num_qtd%>">
		</span>
<%Elseif natureza = "c" then%>
	<input type="radio" name="cd_tipo_item" value="e" onFocus="mostra_itens_c('e');">Equipamento &nbsp; &nbsp; &nbsp; &nbsp; 
	<input type="radio" name="cd_tipo_item" value="o" onFocus="mostra_itens_c('o');">Ótica &nbsp; &nbsp; &nbsp; &nbsp; 
	<input type="radio" name="cd_tipo_item" value="i" onFocus="mostra_itens_c('i');">Instrumento &nbsp; &nbsp; &nbsp; &nbsp; 
	<input type="radio" name="cd_tipo_item" value="m" onFocus="mostra_itens_c('m');">Material
		<br>&nbsp;
		<span id="divFase1">
			<input type="hidden" name="cd_patrimonio" value="<%=cd_patrimonio%>">
			<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
			<input type="hidden" name="cd_marca" value="<%=cd_marca%>">
			<input type="hidden" name="cd_especialidade" value="<%=cd_especialidade%>">
			<input type="hidden" name="cd_ns" value="<%=cd_ns%>">
			<input type="hidden" name="num_qtd" value="<%=num_qtd%>">
		</span>
<%Elseif natureza = "p" then%>
	<input type="hidden" name="cd_equipamento" value="0">
	<input type="hidden" name="num_qtd" value="0">
	>>> Realização de Manutenção Preventiva em Equipamentos <<<
		
<%else%>
<%=natureza%> + <%=cod_os%> ?
<%end if%>