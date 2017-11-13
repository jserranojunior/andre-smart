<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza cargos</title>
</head>
<%cd_funcionario = request("cd_funcionario")
cd_funcao = request("cd_funcao")
%>
<body>
<table cellspacing="1" cellpadding="1" border="0">
<form action="..\..\adm\adm3_acao.asp" method="post" name="funcao" id="funcao">

<input type="hidden" name="acao" value="inserir">
<input type="hidden" name="cd_tipo" value="cargo">
<input type="hidden" name="cd_funcionario" value="<%=cd_funcionario%>">
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
    <td>Cargo atual: <%=cd_funcao%></td>
</tr>
<tr>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Selecione o novo cargo:</td>
</tr>
<% strsql = "TBL_funcoes where cd_especificidade = 3"
  Set rs_func = dbconn.execute(strsql)
  'nm_funcao =  rs_func("nm_funcao")%>
  <%'if strcod = "" Then%>								
<tr>
    <td><select name="cd_funcao" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs_func.EOF
	'rs_func("cd_codigo")%>
	<option value="<%=rs_func("cd_codigo")%>"<%if strcd_funcao = CInt(rs_func("cd_codigo")) then%>selected<%End if%>><%=rs_func("nm_funcao")%></option>
	  <%rs_func.movenext
		loop

		'rs_func.close
		'Set rs_func = nothing
	  %></td>
</tr>
<tr>
    <td><select name="cd_dia" class="inputs">
							<%dia_data = 1
							do while not dia_data > 31%>
							<option value="1"><%=dia_data%></option>
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
    <td><input type="submit" value="Atualiza"></td>
</tr>
</table>


</form>
</body>
</html>
