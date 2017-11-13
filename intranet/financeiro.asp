
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
   <div id="frame">
		<%'if tipo = "diario" Then%>
		<!--include file="financeiro/diario.asp"-->
		<%'Elseif tipo = "diario2" Then%>
		<!--include file="financeiro/diario2.asp"-->
		<%if tipo = "diario3" Then%>
		<!--#include file="financeiro/diario3.asp"-->
		<%elseif tipo = "pagto_semanal" Then%>
		<!--#include file="financeiro/diario_pagto_semanal3.asp"-->
		<%elseif tipo = "emissao_cheques" Then%>
		<!--#include file="financeiro/diario_emissao_cheques3.asp"-->
		<%elseif tipo = "relat_areas" Then%>
		<!--#include file="financeiro/diario_relat_areas3.asp"-->
		<%elseif tipo = "fatura" Then%>
		<!--#include file="financeiro/fatura3.asp"-->
		<%elseif tipo = "consolidado" Then%>
		<!--#include file="financeiro/fatura_consolidado3.asp"-->
		<%end if%>
	</div>                    
</body>
</html>