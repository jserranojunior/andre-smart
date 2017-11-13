
    <!--#include file="topo.asp"-->
	<%tipo = request("tipo")%>
	
   <tr>
		<td width="1%" bgcolor="#BCBCBC" id="no_print"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="13%" valign="top" bgcolor="#999999" class="txt" id="no_print"><!--#include file="menu.asp"--></td>
		<td valign="top" width="86%">
		<%if tipo = "lista" Then%>
		<!--#include file="empresa/funcionario/funcionarios_lista.asp"-->
		<%elseif tipo = "provisorio" Then%>
		<!--#include file="empresa/funcionario/funcionarios_provisorio.asp"-->
		<%elseif tipo = "novo" Then%>
		<!--#include file="empresa/funcionario/funcionarios_cadastro.asp"-->
		<%elseif tipo = "cadastro" Then%>
		<!--#include file="empresa/funcionario/funcionarios_cadastro.asp"-->
		<%elseif tipo = "cartao_ponto" Then%>
		<!--#include file="empresa/funcionario/cartao_print.asp"-->
		<%elseif tipo = "relatorio" Then%>
		<!--#include file="empresa/funcionario/funcionarios_coren_venc.asp"-->
		<%elseif tipo = "ausencia" Then%>
		<!--#include file="empresa/funcionario/ausencias/funcionarios_ausencias.asp"-->
		<%elseif tipo = "cad_ausencia" Then%>
		<!--#include file="empresa/funcionario/ausencias/funcionarios_cad_ausencias.asp"-->	
		<%elseif tipo = "escala" Then%>
		<!--#include file="empresa/escala/funcionarios_escala.asp"-->
		<%elseif tipo = "movimentacao" Then%>
		<!--#include file="empresa/funcionario/funcionarios_movimentacao.asp"-->
		<%elseif tipo = "mov_anual" Then%>
		<!--#include file="empresa/funcionario/funcionarios_movimentacao_anual.asp"-->
		<%elseif tipo = "valor_padrao" Then%>
		<!--#include file="empresa/funcionario/funcionarios_valor_padrao.asp"-->
		
		
		
		<%'elseif tipo = "listagem" Then%>
		<!--include file="empresa/funcionario/funcionarios_coren_lista.asp"-->
		<%elseif tipo = "validade" Then%>
		<!--#include file="empresa/funcionario/funcionarios_coren_venc.asp"-->
		<%elseif tipo = "dependentes" Then%>
		<!--#include file="empresa/funcionario/funcionarios_dependentes_lista.asp"-->
		
		<%elseif tipo = "relatorio_faltas" Then%>
		<!--#include file="relatorios/empresa/funcionarios_faltas.asp"-->
		<%elseif tipo = "vencimento" Then%>
		<!--#include file="empresa/funcionario/funcionarios_coren_venc.asp"-->
		<%elseif tipo = "digitacao_horas" Then%>
		<!--#include file="empresa/horas_trabalhadas/horas_trabalhadas_cartao.asp"-->		
		<%elseif tipo = "encerramento" Then%>
		<!--#include file="empresa/encerramento/empresa_encerramento.asp"-->
		<%end if%>
		</td>
   </tr>                           
   </table>                    
</body>
</html>