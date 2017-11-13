<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<html>
<head>
	<title>Atualiza Status</title>
</head>
<%cd_funcionario = request("cd_funcionario")
cd_status = request("cd_status")
%>
<body>
<table cellspacing="1" cellpadding="1" border="0">
<form action="..\..\acoes\funcionarios_acao.asp" method="post" name="status" id="status">
<input type="hidden" name="acao" value="status">
<input type="hidden" name="cod" value="<%=cd_funcionario%>">
<tr>
    <td>Nome: <%=cd_funcionario%></td>
</tr>
<% strsql = "TBL_funcionario_status where cd_funcionario='"&cd_funcionario&"'  order by cd_codigo Desc"
  Set rs_status = dbconn.execute(strsql)
  if not rs_status.EOF Then
  cd_status = rs_status("cd_status")
  end if
  
  %>
<!--tr>
    <td>Cargo atual: <%=cd_status%></td>
</tr-->
<tr>
    <td>&nbsp;</td>
</tr>
<tr>
    <td>Selecione o novo status:</td>
</tr>
<% strsql = "SELECT * FROM TBL_status where cd_codigo <> 4"
  Set rs = dbconn.execute(strsql)
  'nm_funcao =  rs_func("nm_funcao")%>
  <%'if strcod = "" Then%>								
<tr>
    <td><select name="cd_status" class="inputs"> 
	<option value="">Selecione</option>
	<%Do While Not rs.EOF
	'rs_func("cd_codigo")%>
	<option value="<%=rs("cd_codigo")%>"<%if strcd_status = CInt(rs("cd_codigo")) then%>selected<%End if%>><%=rs("nm_status")%></option>
	  <%rs.movenext
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
