
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	<%tipo = request("tipo")%>
	
	<div id="frame">
		<%if tipo = "villa" Then%>
		<!--include file="agenda/villa.asp"-->
		<%elseif tipo = "s_luiz" Then%>
		<!--include file="agenda/s_luiz.asp"-->
		<%elseif tipo = "agendamento" Then%>
		<!--#include file="agenda/agendamento.asp"-->
		<%end if%>
	</div>                    
</body>
</html>