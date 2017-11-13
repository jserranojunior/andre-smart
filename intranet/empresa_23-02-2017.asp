
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
   <div id="frame">
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
		<%elseif tipo = "aparelho" Then%>
		<!--#include file="empresa/funcionario/aparelho_print.asp"-->
		<%elseif tipo = "relatorio" Then%>
		<!--#include file="empresa/funcionario/funcionarios_coren_venc.asp"-->
		<%elseif tipo = "ausencia" Then%>
		<!--#include file="empresa/funcionario/ausencias/funcionarios_ausencias.asp"-->
		<%elseif tipo = "cad_ausencia" Then%>
		<!--#include file="empresa/funcionario/ausencias/funcionarios_cad_ausencias.asp"-->	
		<%elseif tipo = "escala" Then%>
		<!--#include file="empresa/escala/funcionarios_escala.asp"-->
		
		<%elseif tipo = "folha" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento.asp"-->
		<%elseif tipo = "variaveis" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_var.asp"-->
		
		<%elseif tipo = "relat_vt" Then%>
		<!--include file="empresa/funcionario/funcionarios_folhapagamento_relat_vt.asp"-->
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_var.asp"-->
		<%elseif tipo = "relat_vr" Then%>
		<!--include file="empresa/funcionario/funcionarios_folhapagamento_relat_vt.asp"-->
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_var.asp"-->
		<%elseif tipo = "relat_vt2" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_relat_vt.asp"-->
		<!--include file="empresa/funcionario/funcionarios_folhapagamento_var.asp"-->
		
		<%elseif tipo = "beneficios" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_beneficios.asp"-->
		<%elseif tipo = "ben_extra" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_ben_extra.asp"-->
		
				
		<%elseif tipo = "documentacao" Then%>
		<!--#include file="empresa/funcionario/funcionarios_documentacao.asp"-->
		
		<%elseif tipo = "recibos" Then%>
		<!--#include file="empresa/funcionario/funcionarios_recibos.asp"-->
		
		
		<%elseif tipo = "movimentacao" Then%>
		<!--include file="empresa/funcionario/funcionarios_movimentacao.asp"-->
		<%elseif tipo = "mov_anual" Then%>
		<!--#include file="empresa/funcionario/funcionarios_movimentacao_anual.asp"-->
		<%elseif tipo = "valor_padrao" Then%>
		<!--include file="empresa/funcionario/funcionarios_valor_padrao.asp"-->
			
		
		
		<%'elseif tipo = "listagem" Then%>
		<!--include file="empresa/funcionario/funcionarios_coren_lista.asp"-->
		
		<%elseif tipo = "dependentes" Then%>
		<!--#include file="empresa/funcionario/funcionarios_dependentes_lista.asp"-->
		
		<%elseif tipo = "relatorio_faltas" Then%>
		<!--#include file="relatorios/empresa/funcionarios_faltas.asp"-->
		
		<%elseif tipo = "vencimento" Then%>
		<!--#include file="empresa/funcionario/funcionarios_coren_venc.asp"-->
		<%elseif tipo = "coren_venc" Then%>
		<!--#include file="empresa/funcionario/funcionarios_coren_venc.asp"-->
		<%elseif tipo = "exammed_venc" Then%>
		<!--#include file="empresa/funcionario/funcionarios_exammed_venc.asp"-->
		<%elseif tipo = "dupla_venc" Then%>
		<!--#include file="empresa/funcionario/funcionarios_vacinas_duplo.asp"-->
		<%elseif tipo = "hepatb_vacina" Then%>
		<!--#include file="empresa/funcionario/funcionarios_vacinas_hepatb.asp"-->
		<%elseif tipo = "scr_vacina" Then%>
		<!--#include file="empresa/funcionario/funcionarios_vacinas_scr.asp"-->
		
		
		<%elseif tipo = "digitacao_horas" Then%>
		<!--#include file="empresa/horas_trabalhadas/horas_trabalhadas_cartao.asp"-->		
		<%elseif tipo = "encerramento" Then%>
		<!--#include file="empresa/encerramento/empresa_encerramento.asp"-->
		<%elseif tipo = "resumo_venc" Then%>
		<!--#include file="empresa/funcionario/funcionarios_vencimentos.asp"-->
		<%end if%>
	</div>                    
</body>
</html>