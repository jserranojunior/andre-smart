<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%'txtBusca = request("txtBusca")
cd_user = session("cd_codigo")

busca_aparelho = request("busca_aparelho")

'response.write(busca_aparelho)
busca_aparelho = busca_ajax(busca_aparelho)

	separador1 = instr(1,busca_aparelho,"|",1)
	aparelho = trim(mid(busca_aparelho,1,separador1-1))
	status = mid(busca_aparelho,separador1+1,len(busca_aparelho)-separador1)


num_lista = 1	%>

<%'********************************
'*** Limpa o resultado da busca ***
'**********************************
If  aparelho = "" Then%>
<table border="0" width="500">		
	<%if num_lista = 1 Then%>
				<tr bgcolor="#C0C0C0">
					<td align="center">Nº</td>
					<td>&nbsp;<b>NOME</b></td>
					<td>&nbsp;<b>TIPO</b></td>
					<td>&nbsp;<%=cd_user%></td>
				</tr>
				<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
					<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
					<td><img src="../imagens/px.gif" alt="" width="300" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td>				
				</tr>
	<%end if%>	
</table>
<%'*******************************
'*** Mostra elementos da busca ***
'*********************************
Elseif aparelho = "todos" Then%>
	<%strsql ="SELECT * FROM View_aparelhos_comunicacao WHERE cd_status = "&status&" ORDER BY nr_numero"
		Set rs = dbconn.execute(strsql)
		num_lista = 1%>
	<table border="0" width="500">		
		<%Do while Not rs.EOF
			cd_aparelho = rs("cd_codigo")
			nr_numero = rs("nr_numero")
				if nr_numero <> "" Then
				nr_numero_pref = mid(nr_numero,1,1)
				nr_numero_1 = mid(nr_numero,2,4)
				nr_numero_2 = mid(nr_numero,6,4)
				nr_numero = nr_numero_pref&"-"&nr_numero_1&"-"&nr_numero_2
				end if
			nm_modelo = rs("nm_modelo")
			cd_status = rs("cd_status")%>
	
	<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">
			<td align="center">Nº</td>
			<td>&nbsp;<b>Aparelho</b></td>
			<td>&nbsp;<b>Modelo</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		</tr>
	<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">			
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<!--td align="center" valign="top"><b>&nbsp;<%'=zero(num_lista)%></b><%=cd_status%></td-->
		<td valign="top">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/aparelhos_cad.asp?cd_aparelho=<%=cd_aparelho%>&status=<%=cd_status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;">&nbsp;<%=nr_numero%></td>
		<td valign="top">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/aparelhos_cad.asp?cd_aparelho=<%=cd_aparelho%>&status=<%=cd_status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;">&nbsp;<%=nm_modelo%>-<%=cd_status%></td>
		<!--td valign="top">&nbsp;<%=dt_cadastro%></td-->
		<td id="no_print" valign="top" align="center"><img src="imagens/ic_pendente.gif" onclick="javascript:JsAtivaDesativa('<%=cd_aparelho%>')" type="button" value="A" title="Desativar" class="no_print"></td>		
			<%num_lista = num_lista + 1
			rs.movenext
			loop%>
		
	</tr>
	<!--tr><td>Todos</td></tr-->
</table>
<%'******************
'*** Mostra busca ***
'********************
Else%>
	<%
	if IsNumeric(left(aparelho,1)) then
		aparelho = replace(aparelho,"-","")
		'response.write("++"&aparelho)
		xsql = "SELECT * FROM View_aparelhos_comunicacao where nr_numero like '"&aparelho&"%' and cd_status = "&status&" ORDER BY nr_numero"
	elseif Not IsNumeric(mid(aparelho,1,1)) then
		xsql = "SELECT * FROM View_aparelhos_comunicacao where nm_modelo like '%"&aparelho&"%' and cd_status = "&status&" ORDER BY nm_modelo"
	Else
	 xsql = ""
	end if
	Set rs = dbconn.execute(xsql)%>
	<%num_lista = 1%>
<table border="0" width="500">	
	<%'if not rs.EOF Then%>
		<%while not rs.EOF%>
		<%cd_aparelho = rs("cd_codigo")
		nr_numero = rs("nr_numero")
			if nr_numero <> "" Then
			nr_numero_pref = mid(nr_numero,1,1)
			nr_numero_1 = mid(nr_numero,2,4)
			nr_numero = nr_numero_pref&"-"&nr_numero_1&"-"&mid(nr_numero,6,4)
			end if
		nm_modelo = rs("nm_modelo")
		cd_status = rs("cd_status")%>		
		<%if num_lista = 1 Then%>
		<tr bgcolor="#C0C0C0">		
			<td align="center">Nº</td>
			<td>&nbsp;<b>Aparelho</b></td>
			<td>&nbsp;<b>Modelo</b></td>
			<td>&nbsp;</td>
		</tr>
		<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
			<td><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td>	
			<td><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
			<td><img src="../imagens/px.gif" alt="" width="200" height="1" border="0"></td>
		</tr>
		<%end if%>
	<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'');">
		<td align="right"><%=num_lista%>&nbsp;&nbsp;</td>
		<td>&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/aparelhos_cad.asp?cd_aparelho=<%=cd_aparelho%>&status=<%=status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;"><%=nr_numero%></a></td>
		<td valign="top">&nbsp;<a href="javascript:void(0);" onClick="window.open('ferramentas/aparelhos_cad.asp?cd_aparelho=<%=cd_aparelho%>&status=<%=status%>&acao=editar','Altera_nome','width=500,height=300,scrollbars=1')" return false;">&nbsp;<%=nm_modelo%></td>
		<td id="no_print" valign="top" align="center">
			<img src="imagens/ic_pendente.gif" onclick="javascript:JsAtivaDesativa('')" type="button" value="A" title="Desativar" class="no_print"><%'=cd_status%>
			<!--img src="imagens/ic_pendente.gif" onclick="javascript:JsDesativar('<%=cd_aparelho%>')" type="button" value="A" title="Desativar"--><%'=cd_aparelho%>
		</td>
		<!--td><%'=nm_equipamento_novo%></td-->
	</tr>
		<%num_lista = num_lista + 1
		rs.MoveNext
		wend
	'else%>
		
	<%'End if
	  rs.close
	  set rs = nothing%>	   
</table>
	

<%'else%>
	<!--table>
		<tr>
			<td>Nenhum</td>
		</tr>
	</table-->
<%end if%>
<br>
<span style="color:silver;"><%=fornecedor%></span>

<SCRIPT LANGUAGE="javascript">
/*function JsDesativar(){
alert("teste");
}*/
/*function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar o equipamento inativo?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=0&acao=excluir');
	  }
}*/
</SCRIPT>