
  <!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
   <!--#include file="../includes/inc_area_restrita.asp"-->

<%
cd_menu = request("cd_menu")
cd_menu2 = request("cd_menu2")

item = request("item")
link = request("link")
parametro = request("parametro")
	parametro = replace(parametro,"&","&amp;")
	
TBL_destino = request("tbl_destino")
sub_tipo = request("sub_tipo")

if TBL_destino = "1"  AND sub_tipo = "links" then
xsql = "up_menu_1_links_update @cd_menu='"&cd_menu&"', @link='"&link&"', @parametro='"&parametro&"'"
response.write("menu1-upd")
dbconn.execute(xsql)
Elseif TBL_destino = "2" AND sub_tipo = "links" then
xsql = "up_menu_2_links_update @cd_menu='"&cd_menu&"', @link='"&link&"', @parametro='"&parametro&"'"
response.write("menu2-upd")
dbconn.execute(xsql)
Elseif TBL_destino = "3" AND sub_tipo = "links" then
xsql = "up_menu_3_links_update @cd_menu='"&cd_menu&"', @link='"&link&"', @parametro='"&parametro&"'"
response.write("menu3-upd")
dbconn.execute(xsql)

Elseif TBL_destino = "1"  AND sub_tipo = "itens" then
xsql = "up_menu_1_itens_insert @item='"&item&"', @link='"&link&"', @parametro='"&parametro&"'"
dbconn.execute(xsql)
Elseif TBL_destino = "2"  AND sub_tipo = "itens" then
xsql = "up_menu_2_itens_insert @item='"&item&"', @menu_principal='"&cd_menu&"', @link='"&link&"', @parametro='"&parametro&"'"
dbconn.execute(xsql)
Elseif TBL_destino = "3"  AND sub_tipo = "itens" then
xsql = "up_menu_3_itens_insert @item='"&item&"', @menu_principal='"&cd_menu&"', @menu_medio='"&cd_menu2&"', @link='"&link&"', @parametro='"&parametro&"'"
dbconn.execute(xsql)
Else
response.write("!!!!! ERRO !!!!!<br><br>")
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>
<body onload="window.close()">
<!--body-->
Codigo - <%=cd_menu%><br>
Codigo - <%=cd_menu2%><br>
item = <%=item%><br>

Link = <%=link%><br>
Parametro = <%=parametro%><br>
*TBL_destino = <%=TBL_destino%><br>
*Sub tipo = <%=sub_tipo%><br>

</body>
</html>