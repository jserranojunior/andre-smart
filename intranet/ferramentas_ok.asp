
    <!--#include file="topo.asp"-->
	<%tipo = request("tipo")%>
	
   <tr>
		<td width="1%" bgcolor="#BCBCBC" id="no_print"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="15%" valign="top" bgcolor="#999999" class="txt" id="no_print"><!--#include file="menu.asp"--></td>
		<td valign="top" width="84%">		
		<%'if tipo = "" Then%>
			<!--include file="ferramentas/ferramentas_menu.asp"-->
		<%'else
		if tipo = "medico" then%>
			<!--#include file="ferramentas/medico_cad.asp"-->
		<%elseif tipo = "convenio" then%>
			<!--#include file="ferramentas/convenio_cad.asp"-->
		<%elseif tipo = "rack" then%>
			<!--#include file="ferramentas/rack_cad.asp"-->
		<%elseif tipo = "espec" then%>
			<!--#include file="ferramentas/espec_cad.asp"-->
		<%elseif tipo = "especialidade" then%>
			<!--#include file="ferramentas/especialidade_cad.asp"-->
		<%elseif tipo = "procedimento" then%>
			<!--#include file="ferramentas/procedimento_cad.asp"-->
		<%elseif tipo = "material" then%>
			<!--#include file="ferramentas/material_cad.asp"-->
		<%elseif tipo = "equipamento" then%>
			<!--#include file="ferramentas/equipamento_cad.asp"-->
		<%elseif tipo = "unidade" then%>
			<!--#include file="ferramentas/unidade_cad.asp"-->
		<%end if%>
		</td>
   </tr>                           
   </table>                    
</body>
</html>