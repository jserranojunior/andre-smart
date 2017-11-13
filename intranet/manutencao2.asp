	<%tipo = request("tipo")%>
	
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	
   <div id="frame">
		<%if tipo = "pendencias" Then%>
		<!--#include file="manutencao_2/manutencao_pendencias2.asp"-->
		<%elseif tipo = "especialidades" Then%>
		<!--include file="manutencao/manutencao_pendencias_especialidades.asp"-->
		<%elseif tipo = "andamento" Then%>
		<!--#include file="manutencao_2/manutencao_andamento2.asp"-->
		<%elseif tipo = "andamento3" Then%>
		<!--include file="manutencao_2/manutencao_andamento3.asp"-->
		<%elseif tipo = "nova" Then%>
		<!--include file="manutencao/manutencao_nova.asp"-->
		<%elseif tipo = "nova2" Then%>
		<!--#include file="manutencao_2/manutencao_nova.asp"-->
		<%elseif tipo = "nova3" Then%>
		<!--#include file="manutencao_2/manutencao_nova_3.asp"-->
		<%elseif tipo = "relatorio" Then%>
		<!--#include file="manutencao_2/manutencao_relatorio.asp"-->
		<%elseif tipo = "relatorio2" Then%>
		<!--#include file="manutencao_2/manutencao_relatorio2.asp"-->
		<%elseif tipo = "gestao" Then%>
		<!--#include file="manutencao_2/manutencao_relatorio_gestao.asp"-->
		<%elseif tipo = "periodo" Then%>
		<!--include file="manutencao/manutencao_relatorio_periodo.asp"-->
		<%elseif tipo = "cad_orc_aprov" Then%>
		<!--#include file="manutencao_2/cad_orc_aprov.asp"-->
		<%elseif tipo = "orc_aprov" Then%>
		<!--#include file="manutencao_2/orcamentos_aprovados.asp"-->
		<%elseif tipo = "man_rep_aprov" Then%>
		<!--#include file="manutencao_2/rel_orc_aprov_2.asp"-->
		<%elseif tipo = "rel_orcamentos" Then%>
		<!--#include file="manutencao_2/rel_orc_aprov.asp"-->
		<%end if%>
   </div>                    
</body>
</html>