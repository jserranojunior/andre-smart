
    <!--#include file="topo.asp"-->
   <tr>
		<td width="1%" bgcolor="#BCBCBC" id="no_print"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="15%" valign="top" bgcolor="#999999" class="txt" id="no_print"><!--#include file="menu.asp"--></td>
		<td valign="top" width="84%">
		<%ver = request("ver")
		
		if strsession = 2 Then
		'ver = "old"
		'ver = "nova"
		end if
		
		if ver = "nova" Then%>
		<!--#include file="manutencao/manut_os_4.asp"-->
		<%Elseif ver <> "old" Then%>
		<!--#include file="manutencao/manut_os_3.asp"-->
		<%else%>
		<!--#include file="manutencao/manut_os_3.asp"-->
		<%end if%>
		</td>
   </tr>                           
   </table>                    
</body>
</html>