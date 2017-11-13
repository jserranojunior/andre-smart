
    <!--#include file="topo.asp"-->
	<%tipo = request("tipo")%>
	
   <tr>
		<td width="1%" bgcolor="#BCBCBC" id="no_print"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="15%" valign="top" bgcolor="#999999" class="txt" id="no_print"><!--#include file="menu.asp"--></td>
		<td valign="top" width="84%">
		<%if tipo = "lista" Then%>
		<!--#include file="empresa/funcionario/funcionarios_lista.asp"-->
		<%elseif tipo = "novo" Then%>
		<!--#include file="empresa/funcionario/funcionarios_cadastro.asp"-->
		<%elseif tipo = "cadastro" Then%>
		<!--#include file="empresa/funcionario/funcionarios_cadastro.asp"-->
		<%end if%>
		</td>
   </tr>                           
   </table>                    
</body>
</html>