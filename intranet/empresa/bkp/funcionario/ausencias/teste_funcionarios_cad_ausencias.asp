<!--#include file="../../../includes/util.asp"-->
<!--#include file="../../../includes/inc_open_connection.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<%   
strcod = request("cod")
	if strcod = "" Then
		strcod = request("id")
	end if
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

If strcod <> "" And strmsg = "" Then
	
xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
Set rs = dbconn.execute(xsql)

	If NOT rs.EOF Then
		str_funcionario = rs("cd_codigo")
		strregime_trabalho = rs("cd_regime_trabalho")
		strmatricula = rs("cd_matricula") 
		strnome = rs("nm_nome")
		str_sexo = rs("cd_sexo")
		strfoto = rs("nm_foto")
		strregime_contrato = rs("cd_regime_trabalho")
		
		cd_numreg = rs("cd_numreg")
		str_numreg = rs("cd_numreg")
	End If

else

End if
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/padrao.js"></SCRIPT>
<!--include file="../../../js/ajax.js"-->
<!--include file="js/cid.js"-->

					<% If strmsg <> "" then%>
						<table  border="0" cellspacing="0" cellpadding="0" class="txt_cinza">
							<tr id="no_print">
							 <td>
								<b><%=strmsg%></b>
							 </td>
							</tr>
							<tr id="no_print">
							 <td align=center><br><br><a href="coop_cooperados_lista.asp"><img src="imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
							</tr>
						</table>
					<%else%>
					<table align="center" border="0" rules="groups" width="100%">
					<tr id="no_print">
						<td>&nbsp;&nbsp;<a href="empresa.asp?tipo=ausencia&func=ativo">Listagem</a> 
						&nbsp;&nbsp;&nbsp; <a href="empresa.asp?tipo=novo">Novo</a></td>
					</tr></table>
					<table align="center" border="1" rules="groups">
					<tr><td>
					<table border="0" cellspacing="0" cellpadding="2" align="center"><!-- OK 2-->
						
						<form name="Form" action="empresa/acoes/funcionarios_acao.asp" method="POST">
						
							<tr>
								<td valign="top">
									<input type="text" name="dt_dia" size="3" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(dia_sel)%>">/
									<input type="text" name="dt_mes" size="3" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=zero(mes_sel)%>">/
									<input type="text" name="dt_ano" size="5" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=ano_sel%>">
								</td>
							</tr>							
						</form>	
						</table>
						</td></tr>
						</table>
						<br>
						<br>
							
					
									 								
							<%End if%>