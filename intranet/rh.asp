
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	<%tipo = request("tipo")%>
	
   <div id="frame">
		<%if tipo = "folha" Then%>
		<!--#include file="rh/folha_pagamento.asp"-->		
		<%end if%>
   </div>                    
</body>
</html>