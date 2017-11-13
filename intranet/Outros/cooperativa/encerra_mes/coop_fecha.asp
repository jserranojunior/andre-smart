<!--include file="../../includes/util.asp"-->
<!--include file="../../includes/inc_open_connection.asp"-->
   <!--include file="includes/inc_area_restrita.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
</HEAD>

<%If acao = "encerra" Then

   xsql = "up_Fechamento_mes @cd_mes='"&mes_sel&"', @cd_ano='"&ano_sel&"', @nm_fechamento=true"
	dbconn.execute(xsql)
	strmsg = "Benefício cadastrado com sucesso!"
 End If%>
<LINK href="css/estilo.css " type=text/css rel=stylesheet >
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>

<BODY> 

<%
xsql = "Select * From TBL_fechamento_mes Order by cd_ano,cd_mes"
Set rs = dbconn.execute(xsql)

'Verifica o ultimo mes fechado
Do While Not rs.EOF

strmes_fechado = rs("cd_mes")
strano_fechado = rs("cd_ano")

'strmes_hoje = strmes + 1
'	if strmes + 1 = 13 Then
'	strmes_now = 1
'	strano_now = strano + 1
'	end if
	
strmes_hoje = Month(now)
strano_hoje = Year(now)

	if strmes_hoje = 1 Then
	strmes_nome = strmes_hoje
	strmes_hoje = 13
	'strano_hoje = strano + 1
	end if
	
rs.movenext

loop

	'*** acrescenta 0 ao mes inferior a 10 ***
	if strmes_hoje < 10 Then
	strmes_hoje = 0&strmes_hoje
	end if
	
	if strmes_fechado < 10 Then
	strmes_fechado = 0&strmes_fechado
	end if

	'if strmes_fechado = 12 then
	


'**** Início do uso do sistema ****
if strmes_fechado = "" AND strano_fechado = "" Then
strmes_fechado = 1
strano_fechado = 2007
end if

strdata_fechada = strano_fechado&strmes_fechado
strdata_hoje = strano_hoje&strmes_hoje

%><br>
<table align="center" border="1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
			<table width=350 border=0 cellspacing="0" cellpadding="0">
				<tr>
					<td class="txt_cinza"><b>Relatórios &raquo; <font color="red">Cooperativa</font></b><br><br></td>
				</tr>	
				<tr>
					<td align=center bgcolor=black colspan=3></td>
				</tr>
				<tr rowspan="2">
					<td  align=left class="textopadrao"><br>
					<font color="red">&nbsp;Último mês encerrado <%=strmes_fechado%><%=mes_selecionado(strmes_fechado)%>/<%=strano_fechado%>.</font><br><br></td>
				</tr>
				<tr bgcolor=#cococo class="textopadrao">
					
				 </tr>
				 <tr class="textopadrao">
				 	<td align=left colspan=2></td>
				</tr><!-- Janeiro -->
					<%If strano_fechado + 1 = strano_hoje AND strmes_fechado + 1 = strmes_hoje Then%>
				<tr>
					<td>Mês (<%=mes_selecionado(strmes_nome)%>/<%=strano_hoje%>) ainda está em andamento.<br><br></td>
				</tr><!-- A partir de Fevereiro -->
					<%Elseif strano_fechado = strano_hoje AND strmes_fechado + 1 = strmes_hoje Then%>
				<tr>
					<td>Mês (<%=mes_selecionado(strmes_nome)%>/<%=strano_hoje%>) ainda está em andamento.<br><br></td>
				</tr>
					<%'Elseif strano_fechado < strano_hoje AND strmes_fechado < strmes_hoje then
						Elseif strdata_fechada < strdata_hoje then%>%>
					<form name="form" action="cooperativa/encerra_mes/coop_fecha_acao.asp" id="beneficios" method="post">
						<input type="hidden" name="acao" value="encerra">
						<%if strmes_fechado + 1 = 13 Then%><%strmes_fechado = 1%><%strmes_fechado = strmes_fechado +1%><%end if%>
						<input type="hidden" name="mes_sel" value="<%=strmes_fechado + 1%>">
						<input type="hidden" name="ano_sel" value="<%=strano_fechado%>">
				
						<input type="hidden" name="teste" value="andre">
				<tr>
					<td align=Center colspan=3><br>
				  	Deseja encerrar <%=mes_selecionado(strmes_fechado+1)%><%'=strmes_fechado%>/<%=strano_fechado%>?<br><br>
				    	<input type="submit" value="OK"><input type="button" value="Voltar" onClick="Javascript:window.location='beneficios_lista.asp?mes=<%=strmes_sel%>'">&nbsp;&nbsp;&nbsp;</td> 
				    
				</tr>
					</form>
					<%Else%>
				<tr>
					<td>Opções não se enquadram.<br><br></td>
				</tr>
				<tr>
					<td align=left colspan=3></td>
				</tr>
				 <%End If%>
				<tr>
					<td align=center bgcolor=white colspan=3><%'=strdata_hoje%> <%'=strdata_fechada%><br>
															 <%'=strmes_fechado%> <%'=strano_fechado%><br>
															 <%'=strmes_hoje%> <%'=strano_hoje%></td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<Br><Br><Br><Br>



</BODY>
</HTML>