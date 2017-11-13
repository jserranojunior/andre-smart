	
    <!--#include file="topo.asp"-->
	<%tipo = request("tipo")%>
	
   <tr>
		<td width="1%" bgcolor="#BCBCBC" id="no_print"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="13%" valign="top" bgcolor="#999999" class="txt" id="no_print"><!--#include file="menu.asp"--></td>
		<td valign="top" width="86%">
		<%if tipo = "digitacao" Then%>
		<!--#include file="protocolo/protocolo_digitacao.asp"-->
		<%elseif tipo = "ver" Then%>
		<!--#include file="protocolo/protocolo_digitacao.asp"-->
		<%elseif tipo = "dvd" Then%>
		<!--#include file="protocolo/protocolo_digitacao.asp"-->
		<%elseif tipo = "dvd_relacao" Then%>
		<!--#include file="protocolo/protocolo_dvd_relacao.asp"-->
		<%elseif tipo = "relatorio" Then%>
		<!--#include file="protocolo/protocolo_relatorio.asp"-->
		<%elseif tipo = "etiquetas" Then%>
		<!--#include file="protocolo/etiquetas/form.asp"-->
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
		<%elseif tipo = "patrimonio" Then%>
		<!--#include file="protocolo/protocolo_relatorio_utilizacao_oticas.asp"-->
		<%elseif tipo = "top10" Then%>
		<!--#include file="protocolo/protocolo_relatorio_top10.asp"-->				
		<%end if%>
		</td>
   </tr>                           
   </table>                    
</body>
</html>