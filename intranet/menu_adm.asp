	<!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	
	<%'tipo = request("tipo")%>
	<div id="frame">	
   		<%if tipo = "lista" Then%>
			<!--#include file="sistema_adm/menu/menu_lista.asp"-->
		<%elseif tipo = "novo" OR tipo = "editar" then%>
			<!--#include file="sistema_adm/usuarios/usuarios_novo.asp"-->
		<%end if%>
   </div>
</body>
</html>