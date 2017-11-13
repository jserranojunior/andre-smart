	
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	<%tipo = request("tipo")%>
	
   <div id="frame">
		<%if tipo = "digitacao" Then%>
		<!--#include file="protocolo/protocolo_digitacao.asp"-->
		<%elseif tipo = "ver" Then%>
		<!--#include file="protocolo/protocolo_digitacao.asp"-->
		<%elseif tipo = "dvd" Then%>
		<!--#include file="protocolo/protocolo_digitacao.asp"-->
		<%elseif tipo = "dvd_relacao" Then%>
		<!--#include file="protocolo/protocolo_dvd_relacao.asp"-->
		<%elseif tipo = "etiquetas" Then%>
		<!--#include file="protocolo/etiquetas/form.asp"-->
		
		<%elseif tipo = "relatorio" Then%>
		<!--#include file="protocolo/protocolo_relatorio.asp"-->
		<%elseif tipo = "med_cirurgias" Then%>
		<!--#include file="protocolo/protocolo_relatorio_medico_proc.asp"-->
				
		<%elseif tipo = "estatistica" Then%>
		<!--#include file="protocolo/protocolo_relatorio_estatistica.asp"-->
		<%elseif tipo = "procedimentos" Then%>
		<!--#include file="protocolo/protocolo_relatorio_procedimentos.asp"-->
		<%elseif tipo = "qtd_digitados" Then%>
		<!--#include file="protocolo/protocolo_relatorio_digitados.asp"-->
		<%elseif tipo = "pendentes" Then%>
		<!--#include file="protocolo/protocolo_relatorio_pendentes.asp"-->
		<%elseif tipo = "emprestimos" Then%>
		<!--#include file="protocolo/protocolo_relatorio_emprestimos.asp"-->
		<%elseif tipo = "otica" Then%>
		<!--#include file="protocolo/protocolo_relatorio_patrimonio_uso.asp"-->
		<%elseif tipo = "top10" Then%>
		<!--#include file="protocolo/protocolo_relatorio_top10.asp"-->
		<%elseif tipo = "mapa_util_racks" Then%>
		<!--#include file="protocolo/protocolo_mapa_util_racks.asp"-->
		<%elseif tipo = "patrimonio" Then%>
		<!--#include file="protocolo/protocolo_relatorio_patrimonio_uso.asp"-->
		<%end if%>              
   </div>                    
</body>
</html>