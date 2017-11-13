
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	
   <div id="frame">
		<%if tipo = "cadastro" Then%>
		<!--#include file="instrumento/instrumento_cad.asp"-->
		<%elseif tipo = "descarte" Then%>
		<!--include file="instrumento/instrumento_desc.asp"-->		
		
		<%elseif tipo = "busca" Then%>
		<!--include file="instrumento/instrumento_busca.asp"-->
		
		<%end if%>
   </div>
</body>
</html>