<!--#include file="../../includes/inc_area_restrita.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%'cd_codigo_usr = request("cd_codigo")
cd_cod_usuario = request("cd_cod_usuario")
msg = request("msg")
%>

<html>


	
<head>
	<title>Untitled</title>
</head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>
<body>
 <SCRIPT LANGUAGE="javascript">
    {       
       objVld=getValidator('form');       
	   objVld.addRule('nm_nome',['required','texto'],'"Dia Entrada"');

		
   }
</SCRIPT>

<%if cd_cod_usuario <> "" Then
'	xsql ="up_adm_lista @cd_codigo='"&cd_cod_usuario&"'"
'	Set rs = dbconn.execute(xsql)
	xsql1 = "SELECT * FROM TBL_ADM_usuario Where cd_codigo='"&cd_cod_usuario&"'"
	Set rs = dbconn.execute(xsql1)
%>
<%if not rs.EOF then
cd_codigo = rs("cd_codigo")
nm_nome_usr = rs("nm_nome")
nm_email_usr = rs("nm_email")
nm_senha_usr = rs("nm_senha")
cd_status = rs("cd_status")
'nm_email = decrypt(rs("nm_email"))
'nm_senha = decrypt(rs("nm_senha"))

'rs.movenext
'loop
end if
End if
%>

<%if cd_cod_usuario = "" Then
acao = "1"
Else
acao = "2"
End if%>

<table width="400" cellspacing="2" cellpadding="2" border="1" class="textopadrao">
<tr bgcolor="#dfdfdf">
    <td colspan="100%"><b>Administração de usuários &raquo; - <font color="red">Cadastro <%'=acao%></font></b></td>
</tr>
<tr>
    <td colspan="100%">&nbsp;<%=cd_cod_usuario%>/<%=acao%></td>
</tr>
		<form action="acoes/usuarios_acao.asp" method="post" name="permissoes" id="permissoes">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
		<input type="hidden" name="cd_cod_usuario" value="<%=cd_cod_usuario%>">
<tr>
    <td>Nome</td>
    <%if cd_cod_usuario = "" then%><td><input type="text" name="nm_nome" size="35" maxlength="30" value=""></td>
	<%else%><td><%=nm_nome_usr%></td><%end if%>
</tr>
<tr>
    <td>E-mail</td>
    <%if cd_cod_usuario = "" then%><td><input type="text" name="nm_email" size="35" maxlength="100" value=""></td>
	<%else%><td><%=nm_email_usr%></td><%end if%>
</tr>
<tr>
    <td>Senha</td>
    <td><input type="password" name="nm_senha" size="35" maxlength="100" value="<%=nm_senha_usr%>"></td>
</tr>
<tr>
    <td>Situação</td>
    <td><select name="cd_status">
			<option value="1" <%if cd_status = 1 then response.write("SELECTED")%>>Ativo</option>
			<option value="0" <%if cd_status = 0 then response.write("SELECTED")%>>Inativo</option>
		</select>
	</td>
</tr>
<tr><td colspan="2"><a href="usuarios.asp?tipo=lista">Listagem</a></td></tr>
</table>
<br>
			<table border="1" cellpadding="0" cellspacing="0">
			<tr bgcolor="#dfdfdf">
    			<td colspan="100%" align="center"><b>Módulos permitidos</b></td>
			</tr>
			<%xsql_upermissoes ="SELECT * FROM tbl_ADM_usuario_permissoes where cd_usuario='"&cd_cod_usuario&"'"
			Set rs_upermissoes = dbconn.execute(xsql_upermissoes)
			if not rs_upermissoes.EOF Then
			permissoes_1 = rs_upermissoes("menu_1")
			permissoes_2 = rs_upermissoes("menu_2")
			permissoes_3 = rs_upermissoes("menu_3")
			End if%>
			
			<%'*** Item principal do menu ***
			xsql_m1 ="SELECT * FROM tbl_menu_1 "
			Set rs_menu_1 = dbconn.execute(xsql_m1)%>			
					<%do while not rs_menu_1.EOF
					cd_menu_1 = rs_menu_1("cd_codigo")
					item_menu_1 = rs_menu_1("item_menu")
					link_menu_1 = rs_menu_1("link_menu")%>
					<tr><%perm_1 = instr(1,permissoes_1,cd_menu_1,0)%><%if perm_1 <> 0 then%><%ck_perm1="checked"%><%else%><%ck_perm1=""%><%end if%>					
						<td valign="top" width="25%"><input type="checkbox" name="menu_1" value="<%=cd_menu_1%>" <%=ck_perm1%>><%=cd_menu_1%> - <%=item_menu_1%><br><font color="#c0c0c0"><%'=link_menu_1%></font></td>
						<td width="75%"><%'*** Item secundário do menu ***
						xsql_m2 ="SELECT * FROM tbl_menu_2 Where menu_principal = "&cd_menu_1&""
						Set rs_menu_2 = dbconn.execute(xsql_m2)%>
						<%'if Not rs_menu_2.EOF Then%>
							<table border="1" width="400">
							<%do while not rs_menu_2.EOF
							cd_menu_2 = rs_menu_2("cd_codigo")
							item_menu_2 = rs_menu_2("item_menu")
							link_menu_2 = rs_menu_2("link_menu")%>
							<tr><%perm_2 = instr(1,permissoes_2,cd_menu_2,0)%><%if perm_2 <> 0 then%><%ck_perm2="checked"%><%else%><%ck_perm2=""%><%end if%>	
								<td valign="top" width="50%"><input type="checkbox" name="menu_2" value="<%=cd_menu_2%>" <%=ck_perm2%>><%=cd_menu_1%>.<%=cd_menu_2%> - <%=item_menu_2%><br><font color="#c0c0c0"><%'=link_menu_2%></font></td>
								<td width="50%"><%'*** Item terciário do menu ***
									xsql_m3 ="SELECT * FROM tbl_menu_3 Where menu_principal = "&cd_menu_1&" AND menu_medio = "&cd_menu_2&""
									Set rs_menu_3 = dbconn.execute(xsql_m3)%>
									<%do while not rs_menu_3.EOF
									cd_menu_3 = rs_menu_3("cd_codigo")
									item_menu_3 = rs_menu_3("item_menu")
									link_menu_3 = rs_menu_3("link_menu")%>
									<%perm_3 = instr(1,permissoes_3,cd_menu_3,0)%><%if perm_3 <> 0 then%><%ck_perm3="checked"%><%else%><%ck_perm3=""%><%end if%>
									<input type="checkbox" name="menu_3" value="<%=cd_menu_3%>" <%=ck_perm3%>><%=cd_menu_1%>.<%=cd_menu_2%>.<%=cd_menu_3%> - <%=item_menu_3%><br><font color="#c0c0c0"><%'=link_menu_3%></font>
									
										<%rs_menu_3.movenext
										loop%>
									&nbsp;
								</td>
							</tr>
							<%'end if
							rs_menu_2.movenext
							loop%>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
					<%'end if
					rs_menu_1.movenext
					loop%>
					<tr><td><input type="submit" name="Ok" value="Ok"></td><td align="right">&nbsp;<img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_codigo%>')" type="button" value="A" title="Apagar"></td></tr>
				</form>				
			</table><br>
<%'response.write(modulos)%>
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acoes/usuarios_acao.asp?cd_cod_usuario='+cod1+'&acao=3');
	  }
}

</SCRIPT>
</body>
</html>




