
    <!--#include file="topo.asp"-->
	<%tipo = request("tipo")%>
	
   <tr>
		<td width="1%" bgcolor="#BCBCBC" id="no_print"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="13%" valign="top" bgcolor="#999999" class="txt" id="no_print"><!--#include file="menu.asp"--></td>
		<td valign="top" width="86%">
		<%if tipo = "cadastro" Then%>
		<!--#include file="patrimonio/patrimonio_cad.asp"-->
		<%elseif tipo = "descarte" Then%>
		<!--#include file="patrimonio/patrimonio_desc.asp"-->
		
		<%elseif tipo = "relat_int" Then%>
		<!--#include file="patrimonio/patrimonio_relat.asp"-->
		<%elseif tipo = "invent_vdlap" Then%>
		<!--#include file="patrimonio/patrimonio_invent_vdlap.asp"-->
		<%elseif tipo = "invent_hosp" Then%>
		<!--#include file="patrimonio/patrimonio_invent_hosp.asp"-->
		<%elseif tipo = "plan_serv" Then%>
		<!--#include file="patrimonio/patrimonio_plan_serv.asp"-->
				
		<%elseif tipo = "relat_otica" Then%>
		<!--#include file="patrimonio/otica_relat.asp"-->
		<%elseif tipo = "oticas_vdlap" Then%>
		<!--#include file="patrimonio/patrimonio_oticas_vdlap.asp"-->
		
		<%elseif tipo = "busca" Then%>
		<!--#include file="patrimonio/patrimonio_busca.asp"-->
		
		
		<%end if%>
		</td>
   </tr>                           
   </table>                    
</body>
</html>