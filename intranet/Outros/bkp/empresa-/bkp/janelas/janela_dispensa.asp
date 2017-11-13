<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza cargos</title>
</head>
<%cd_funcionario = request("cd_funcionario")
cd_dia = request("cd_dia")&"/"
cd_mes = request("cd_mes")&"/"
cd_ano = request("cd_ano")
	dt_desliga = cd_mes&cd_dia&cd_ano

	if dt_desliga <> "//" Then
	response.write("É")
	end if
%>
<body>
<table cellspacing="1" cellpadding="1" border="0">
<form action="..\funcionarios\empr_funcionarios_insert.asp" method="post" name="funcao" id="funcao">
<!--form action="janela_dispensa.asp" method="post" name="funcao" id="funcao"-->
<input type="hidden" name="acao" value="altera">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<input type="hidden" name="cd_status" value="4">
<tr>
    <td>Nome: <%=cd_funcionario%></td>
</tr>
<% strsql = "TBL_funcionario_funcao where cd_funcionario='"&cd_funcaionario&"' order by cd_codigo Desc"
  Set rs_func = dbconn.execute(strsql)
  if not rs_func.EOF Then
  cd_funcao = rs_func("cd_funcao")
  end if
  
  %>
<tr>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Selecione a data da dispensa:</td>
</tr>
<tr>
    <td><select name="cd_dia" class="inputs">
							<%dia_data = 1
							do while not dia_data > 31%>
							<option value="<%=dia_data%>"><%=dia_data%></option>
							<%dia_data = dia_data + 1
							loop%>
						</select>
						<select name="cd_mes" class="inputs">
							<%mes_data = 1
							do while not mes_data > 12%>
							<option value="<%=mes_data%>"><%=mes_selecionado(mes_data)%></option>
							<%mes_data = mes_data + 1
							loop
							%>
						</select>
						<input type="text" name="cd_ano" size="5" maxlength="4" class="inputs" value="<%=year(now)%>"></td>
</tr>
<tr>
    <td><br><input type="submit" value="Dispensar"></td>
</tr>
</table>


</form>
</body>
</html>