
    <!--#include file="topo.asp"-->
	<%tipo = request("tipo")%>
	
   <tr>
		<td width="1%" bgcolor="#BCBCBC"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="15%" valign="top" bgcolor="#999999" class="txt"><!--#include file="menu.asp"--></td>
		<%if tipo = "lista" Then%>
		<td valign="top" align="center" width="84%"><!--#include file="sistema_adm/menu/menu_lista.asp"--></td>
		<%elseif tipo = "novo" OR tipo = "editar" then%>
		<td valign="top" align="center" width="84%"><!--#include file="sistema_adm/usuarios/usuarios_novo.asp"--></td>
		<%end if%>
   </tr>                           
   </table>                    
</body>
</html>