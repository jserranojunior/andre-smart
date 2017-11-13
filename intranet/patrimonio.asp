
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	
   <div id="frame">
		<%if tipo = "cadastro" Then%>
		<!--#include file="patrimonio/patrimonio_cad.asp"-->
		<%elseif tipo = "descarte" Then%>
		<!--#include file="patrimonio/patrimonio_desc.asp"-->
		
		<%elseif tipo = "relat_int" Then%>
		<!--#include file="patrimonio/patrimonio_relat.asp"-->
		<%elseif tipo = "invent_hosp" Then%>
		<!--#include file="patrimonio/patrimonio_invent_hosp.asp"-->
		<%elseif tipo = "plan_serv" Then%>
		<!--#include file="patrimonio/patrimonio_plan_serv.asp"-->
		
		<%elseif tipo = "invent_vdlap" Then%>
		<!--#include file="patrimonio/patrimonio_invent_vdlap.asp"-->
		<%'elseif tipo = "relat_otica" Then%>
		<!--include file="patrimonio/patrimonio_invent_vdlap.asp"-->
		
		<%elseif tipo = "relat_otica" Then%>
		<!--#include file="patrimonio/otica_relat.asp"-->
		<%elseif tipo = "oticas_vdlap" Then%>
		<!--#include file="patrimonio/patrimonio_oticas_vdlap.asp"-->
		<%elseif tipo = "serv_vencer" Then%>
		<!--#include file="patrimonio/patrimonio_serv_vencimento.asp"-->
		<%elseif tipo = "serv_confirmadas" Then%>
		<!--#include file="patrimonio/patrimonio_serv_vencimento.asp"-->
		
		<%elseif tipo = "busca" Then%>
		<!--#include file="patrimonio/patrimonio_busca.asp"-->
		
		<%end if%>
   </div>
</body>
</html>