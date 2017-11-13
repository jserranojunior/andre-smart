<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->

<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<%
'cd_codigo = request("cd_codigo")
campo = request("campo")
list = request("list")
action = request("action")
acao = "inserir"
%><br>

<table width="580" border="1" cellspacing="0" cellpadding="0" align=center>
	<tr>
		<td valign="top" align=center>
			<table width="400" border="1" cellspacing="0" cellpadding="1" class="textopadrao">
				<tr>
					<td class="txt_cinza" colspan="100%"><b>Admin &raquo; - <font color="red">Cadastro de cargos</font></b></td>
				</tr>
					<form method="post" action="adm/adm2_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
					<input type="hidden" name="acao" value="<%=acao%>">
					<input type="hidden" name="tipo" value="cargo">
					<input type="hidden" name="cd_especificidade" value="3">
				<tr bgcolor="#b4b4b4"><td colspan="5" align="center" class="textopadrao">&nbsp;</td></tr>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr bgcolor="#f5f5f5">	
					<td  align="left">Cargo</td>
					<td><input type="text" name="nm_cargo" size="45" maxlength="50" class="inputs" value=""></td>
				</tr>
				<tr bgcolor="#f5f5f5">	
					<td  align="left">Salário</td>
					<td><input type="text" name="cd_salario" size="20" maxlength="20" class="inputs" value=""></td>
				</tr>
				<tr bgcolor="#f5f5f5">	
					<td  align="left">data início</td>
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
						<input type="text" name="cd_ano" size="5" maxlength="4" class="inputs" value=""></td>
				</tr>
				<tr><td colspan="5" align="center"><br>
						<input type="submit" value="inserir cargo">
						<br>
						<br>
						<br>
						<br>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td><b>** CARGOS CADASTRADOS **<br></b></td>
					<td><b>R$</b></td>
				</tr>
				<tr><td colspan="3">&nbsp;</td></tr>
				<%
				selecao = " top 100 percent * "
				tabela = " TBL_funcoes"
				condicoes = " where cd_especificidade = 3"
				criterios = " Order by nm_funcao"
								
				strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
				Set rs = dbconn.execute(strsql)%>
				<%Do while Not rs.EOF
				cd_funcao = rs("cd_codigo")
				nm_funcao = rs("nm_funcao")%>
				<tr>
					<td><!--input type="checkbox" name="cargo" value=""-->&nbsp;</td>
					<td><%=nm_funcao%> - <%=cd_funcao%></td>
					<%xsql_valor="SELECT * FROM TBL_empresa_valores where cd_funcao="&cd_funcao&""
					Set rs_valor = dbconn.execute(xsql_valor)%>
					<td><%=rs_valor("cd_valor")%></td>
				</tr>
				<%rs.movenext
				Loop%>
			</table>
		</td>
	</tr>
</table>												

						