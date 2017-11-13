<%
cd_menu = request("cd_menu")
cd_menu2 = request("cd_menu2")

item_menu1 = request("item_menu1")
item_menu2 = request("item_menu2")
item_menu3 = request("item_menu3")

link = request("link")
parametro = request("parametro")
tbl_destino = request("destino")


%>

<html>
<head>
	<title>Menu - novos itens</title>
</head>

<body onUnload="window.opener.location.reload()">
<table cellspacing="3" cellpadding="2" border="0" align="center">
<form action="../../acoes/menu_acao.asp" name="menu_itens" id="menu_links">
<input type="hidden" name="tbl_destino" value="<%=tbl_destino%>">
<input type="hidden" name="cd_menu" value="<%=cd_menu%>">
<input type="hidden" name="cd_menu2" value="<%=cd_menu2%>">
<input type="hidden" name="sub_tipo" value="itens">
<tr bgcolor="#808080">
    <td colspan="100%" align="center"><b>INSERÇÃO DE ITENS DO MENU</b></td>
</tr>
<%if tbl_destino > 1 Then%>
<tr>
	<td>Principal</td>
	<td><%'=cd_menu&"-"%><%=item_menu1%></td>
</tr>	
<%end if%>
<%if tbl_destino > 2 then%>
	<tr>
		<td>Sub item</td>
		<td><%'=cd_menu2&"-"%><%=item_menu2%></td>
	</tr>
<%end if%>
<tr>
	<td>Item&nbsp;</td>
	<td><input type="text" name="item" value="" maxlength="100" size="30"></td>
</tr>
<tr>
	<td>Link&nbsp;</td>
	<td><%nome_pasta = "C:\inetpub\wwwroot\producao\intranet\"
						'C:\inetpub\wwwroot\producao\intranet
								set FSO = server.createObject("Scripting.FileSystemObject") 
								Set pasta = FSO.GetFolder(nome_pasta) 
								Set arquivos = pasta.Files %>
								<select name="link" size="1" class="inputs">
									<option value="" Selected></option>
									<%if link = nome_arquivo then%><%link_ck="selected"%><%else%><%link_ck=""%><%end if%>
									<option value="#" <%=link_ck%>>#</option>
									<%for each nome_arquivo in arquivos
									'nome_arquivo = Replace(nome_arquivo,"E:\Server\desenvolvimento\producao\intranet\","")
									nome_arquivo = Replace(nome_arquivo,nome_pasta,"")
									if link = nome_arquivo then%><%link_ck="selected"%><%else%><%link_ck=""%><%end if%>
									<option value="<%=nome_arquivo%>" <%=link_ck%>><%'=cd_menu_2&"."%><%=nome_arquivo%></option>
									<%next%>
								</select>
	</td>
</tr>
<tr>
	<td>Parametro&nbsp;</td><td><input type="text" name="parametro" value="<%=parametro%>"></td>
</tr>
<tr><td colspan="100%"><br><img src="imagens/blackdot.gif" width="300" height="1"></td></tr>
<tr>
	<td align="left"><input type="submit" value="OK"></td>
	<%if session("cd_codigo") = 46 then%>
		<td align="center"><span style="font-size=8px;color:red;">menu_itens.asp</span></td>
	<%end if%>
</tr>
<!--tr>
	<td>
	Tabela = <%=tbl_destino%><br>
	Menu 1 = <%=cd_menu%><br>
	
	
	</td>
</tr-->
</table>
</form>
</body>
</html>
