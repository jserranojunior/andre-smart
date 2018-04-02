<%cd_user = session("cd_codigo")
pasta_loc = "sistema_adm\menu\"
arquivo_loc = "menu_links.asp"

cd_menu = request("cd_menu")

item_menu = request("item_menu")
link = request("link")
parametro = request("parametro")
	if parametro <> "" Then
		parametro = replace(parametro,".amp.","&")
	end if
tbl_destino = request("destino")


%>

<html>
<head>
	<title>VD LAP</title>
</head>

<body onUnload="window.opener.location.reload()">
<!--#include file="../../includes/arquivo_loc.asp"-->
<table cellspacing="3" cellpadding="2" border="0" align="center">
<form action="../../acoes/menu_acao.asp" name="menu_links" id="menu_links">
<input type="hidden" name="tbl_destino" value="<%=tbl_destino%>">
<input type="hidden" name="cd_menu" value="<%=cd_menu%>">
<input type="hidden" name="sub_tipo" value="links">
<tr bgcolor="#808080">
    <td colspan="100%" align="center"><b>ALTERAÇÃO DOS LINKS DO MENU</b></td>
</tr>
<tr><td colspan="100%">&nbsp;<%=link%></td></tr>
<tr>
	<td>Item&nbsp;</td>
	<td><%=item_menu%></td>
</tr>
<tr>
	<td>Link&nbsp;</td>
	<td><%
	nome_pasta = "C:\inetpub\wwwroot\producao\intranet\"
								set FSO = server.createObject("Scripting.FileSystemObject") 
								Set pasta = FSO.GetFolder(nome_pasta) 
								Set arquivos = pasta.Files %>
								<select name="link" size="1" class="inputs">
									<option value="" Selected></option>
									<%if link = nome_arquivo then%><%link_ck="selected"%><%else%><%link_ck=""%><%end if%>
									<option value="#" <%=link_ck%>>#</option>
									<%for each nome_arquivo in arquivos
									nome_arquivo = Replace(nome_arquivo,nome_pasta,"")
									if link = nome_arquivo then%><%link_ck="selected"%><%else%><%link_ck=""%><%end if%>
									<option value="<%=nome_arquivo%>" <%=link_ck%>><%'=cd_menu_2&"."%><%=nome_arquivo%></option>
									<%next%>
								</select>
	</td>
</tr>
<tr>
	<td>Parametro&nbsp;</td><td><input type="text" name="parametro" value="<%=parametro%>" size="40"></td>
</tr>
<tr><td><input type="submit" value="OK"></td></tr>
</table>
</form>
</body>
</html>
