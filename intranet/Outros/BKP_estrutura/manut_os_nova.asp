
    <!--#include file="topo.asp"-->
   <tr>
		<td width="1%" bgcolor="#BCBCBC"><img src="px.gif" alt="" width="3" height="1" border="0"></td>
		<td width="15%" valign="top" bgcolor="#999999" class="txt"><!--#include file="menu.asp"--></td>
		<td valign="top" width="84%">
		<%ver = request("ver")
		
		if strsession = 2 Then
		'ver = "old"
		'ver = "nova"
		end if
		
		if ver = "nova" Then%>
		<!--#include file="manutencao/manut_os_nova_2.asp"-->
		<%else%>
		<!--#include file="manutencao/manut_os_nova.asp"-->
		<%end if%><br>
		
		</td>
   </tr>                           
   </table>                    
</body>
</html>