
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
   <div id="frame">
		<%if tipo = "var_enf" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_var.asp"-->
		
		<%elseif tipo = "var_enf_old" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_var_old.asp"-->
		
		<%elseif tipo = "qtd_plantoes" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_plantoes.asp"-->
		<%elseif tipo = "rcm" Then%>
		<!--#include file="empresa/funcionario/funcionarios_folhapagamento_rcm.asp"-->
		
		<%elseif tipo = "documentacao" Then%>
		<!--#include file="empresa/funcionario/funcionarios_documentacao.asp"-->
		
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
		
		
		
		
		<%end if%>
	</div>                    
</body>
</html>