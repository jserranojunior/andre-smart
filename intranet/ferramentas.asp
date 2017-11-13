
    <!--#include file="topo.asp"-->
	<!--#include file="menu.asp"-->
	<%tipo = request("tipo")%>
	
   <div id="frame">		
		<%'if tipo = "" Then%>
			<!--include file="ferramentas/ferramentas_menu.asp"-->
		<%'else
		if tipo = "medico" then%>
			<!--#include file="ferramentas/medico_cad.asp"-->
		<%elseif tipo = "listatodosmedicos"	then%>
			<!--include file="ferramentas/medicos.asp"-->
		<%elseif tipo = "convenio" then%>
			<!--include file="ferramentas/convenio_cad.asp"-->
		<%elseif tipo = "rack" then%>
			<!--include file="ferramentas/rack_cad.asp"-->
		<%elseif tipo = "espec" then%>
			<!--include file="ferramentas/espec_cad.asp"-->
		<%elseif tipo = "especialidade" then%>
			<!--include file="ferramentas/especialidade_cad.asp"-->
		<%elseif tipo = "procedimento" then%>
			<!--include file="ferramentas/procedimento_cad.asp"-->
		<%elseif tipo = "material" then%>
			<!--include file="ferramentas/material_cad.asp"-->
		<%elseif tipo = "equipamento" then%>
			<!--#include file="ferramentas/equipamento_lista.asp"-->
		<%elseif tipo = "unidade" then%>
			<!--#include file="ferramentas/unidade_lista.asp"-->
		<%elseif tipo = "empresa" then%>
			<!--#include file="ferramentas/empresa_lista.asp"-->
		<%elseif tipo = "fornecedores" then%>
			<!--#include file="ferramentas/fornecedores_lista.asp"-->
		<%elseif tipo = "aparelhos" then%>
			<!--#include file="ferramentas/aparelhos_lista.asp"-->
		<%end if%>                          
   </div>                    
</body>
</html>