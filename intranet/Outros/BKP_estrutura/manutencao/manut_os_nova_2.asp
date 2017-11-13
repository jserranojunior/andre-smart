<script language="JavaScript" type="text/javascript" src="../js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../js/mascarainput.js"></SCRIPT>
<script language="javascript">
<!--
  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }

// -->	
</script>
<script language="JavaScript">
function validar_os(shipinfo){
	if (shipinfo.dt_os.value.length==""){window.alert ("A data de abertura da O.S. não foi preenchida.");shipinfo.dt_os.focus();return (false);
	} else {
	if (shipinfo.nm_solicitante.value.length==""){window.alert ("O solicitante não foi informado.");shipinfo.nm_solicitante.focus();return (false);
	} else {
	if (shipinfo.nm_unidade.value.length==""){window.alert ("Selecione uma unidade.");shipinfo.nm_unidade.focus();return (false);
	} else {
	if (shipinfo.num_qtd.value.length==""){window.alert ("Informe a quantidade.");shipinfo.num_qtd.focus();return (false);
	} else {
	if (shipinfo.cd_especialidade.value.length==""){window.alert ("Selecione uma especialidade.");shipinfo.cd_especialidade.focus();return (false);
	} else {
	if (shipinfo.cd_equipamento.value.length==""){window.alert ("Selecione um equipamento.");shipinfo.cd_equipamento.focus();return (false);
	} else {
	if (shipinfo.cd_marca.value.length==""){window.alert ("Selecione uma marca.");shipinfo.cd_marca.focus();return (false);
	} else {
	if (shipinfo.motivo.value.length==""){window.alert ("Qual motivo da solicitação?.");shipinfo.motivo.focus();return (false);
	} else {
	if (shipinfo.dt_entrada.value.length==""){window.alert ("informe a data da providencia.");shipinfo.dt_entrada.focus();return (false);
	} else {
	if (shipinfo.cd_avaliacao.value.length==""){window.alert ("informe o andamento da providencia.");shipinfo.cd_avaliacao.focus();return (false);
	} else {
	if (shipinfo.cd_fornecedor.value.length==""){window.alert ("informe o fornecedor.");shipinfo.cd_fornecedor.focus();return (false);
	} else {
	}}}}}}}}}}}return (true);}
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function foco(){
document.form.dt_dia.focus(); }

nextfield = "dt_dia"; // nome do primeiro campo do site
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln) == "(") break;
	netscape = (ver.charAt(iln+1).toUpperCase() != "C");
	function keyDown(DnEvents) {
		// ve quando e o netscape ou IE
		k = (netscape) ? DnEvents.which : window.event.keyCode;
		if (k == 13) { // preciona tecla enter
		if (nextfield == 'done') {
			//alert("viu como funciona?");
			eval('document.form.' + nextfield + '.focus()');
			return false;
			//return true; // envia quando termina os campos
		} else {
			// se existem mais campos vai para o proximo
			eval('document.form.' + nextfield + '.focus()');
			return false;
		}
	}
}
document.onkeydown = keyDown; // work together to analyze keystrokes
if (netscape) document.captureEvents(Event.KEYDOWN|Event.KEYUP);
// End -->
</script>
<%esconder = "1"%>
<BODY onload="foco()"> 



<br>
<table align="center" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<table width="600" border="1" cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
				<tr>		
					<td class="txt_cinza" colspan="6">
					<b>Manutenção &raquo; - <font color="red">Listagem</font></b><br><br></td>					
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<form action="manut_os_nova.asp" name="nova_os" id="nova_os">
				<tr><td colspan="100%">
					NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok">
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="manutencao.asp?strtop=20&tipo=pendentes">Pendentes</a> 
					/ <a href="ordemservico_listagem.asp?strtop=20&tipo=encerradas">Encerradas</a>
					/ <a href="ordemservico_listagem.asp?strtop=20&tipo=todas">Todas</a>
				</td></tr>
				</form>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<tr>
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						&nbsp;
					</td>
				</tr>
				<%
				
				cd_codigo = request("cd_codigo")
				num_os = request("num_os")
				campo = request("campo")
				ordem = request("ordem")
				list = request("list")
				lista = "1"
			
				
				If campo= "" AND cd_codigo = "" Then
				campo = "num_os"
				formulario = "0"
				End If
			
				compare = "="
				strtop = "100 percent"
				outros = ""
				acao = "1"
				tipo = "cadastra_os"
				
				
'*************************************
'*** Verifica a existência da O.S. ***
'*************************************
if num_os <> "" Then
	xsql = "SELECT * FROM TBL_OrdemServico where num_os="&num_os&""
	Set rs_verifica = dbconn.execute(xsql)
		do while not rs_verifica.EOF
		num_os_verifica = rs_verifica("num_os")
		cd_codigo_verifica = rs_verifica("cd_codigo")
		rs_verifica.movenext
		loop
			if num_os_verifica = int(num_os) then 'Redireciona à O.S. existente
				response.redirect("manut_os_andamento.asp?cd_codigo="&cd_codigo_verifica&"&campo=cd_codigo")
			End if
end if
'*************************************



				
				if cd_codigo <> "" Then
				xsql ="up_os_lista @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
				Set rs = dbconn.execute(xsql)
				
				num_os = rs("num_os")
				dt_os = rs("dt_os")
				nm_unidade = rs("nm_unidade")
				nm_especialidade = rs("nm_especialidade")
				nm_equipamento = rs("nm_equipamento")
				cd_ns = rs("cd_ns")
				cd_patrimonio = rs("cd_patrimonio")
			
				dt_entrada = rs("dt_entrada")
				cd_avaliacao = rs("cd_avaliacao")
				nm_fornecedor = rs("nm_fornecedor")
				cd_fornecedor = rs("cd_fornecedor")
				cd_liberacao = rs("cd_liberacao")
			
				cd_codigo = rs("cd_codigo")
			
				num_qtd = rs("num_qtd")
				nm_marca = rs("nm_marca")
				motivo = rs("motivo")
				nm_solicitante = rs("nm_solicitante")
				observacoes = rs("observacoes")
				
				acao = "2"
				tipo = "modifica_os"
				sequencia = "1"
				fecha = rs("fecha")
				End if%>
			<!--Início da etapa 1.1-->
			<!--Início do formulário-->
				<tr>			
					<form method="post" action="manutencao/manut_os_acao.asp" name="form"  id="form">
						<input type="hidden" name="acao" value="<%=acao%>">
						<input type="hidden" name="tipo" value="<%=tipo%>">
						<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
						<input type="hidden" name="sequencia" value="1">
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
				<tr bgcolor="#b4b4b4">
					<td  colspan="100%" align="center" class="textopadrao">SOLICITAÇÃO</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">O.S.</td>
					<td colspan="4"><%=num_os%>&nbsp;</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Data</td>
					<td colspan="4">
					<input type="text" name="dt_dia" value="<%=dt_dia%>" size="2" maxlength="2" onFocus="nextfield ='dt_mes';" class="inputs">
					<input type="text" name="dt_mes" value="<%=dt_mes%>" size="2" maxlength="2" onFocus="nextfield ='dt_ano';" class="inputs">
					<input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4" onFocus="nextfield ='nm_solicitante';" class="inputs">
					</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Solicitante</td>
					<td colspan="4"><input type="text" name="nm_solicitante" size="40" maxlength="40" class="inputs" value="<%=nm_solicitante%>" onFocus="nextfield ='nm_unidade';"></td>							
				<tr>
					<td align="center" bgcolor="#f5f5f5">Unidade</td>
						<%'****** Mostra as unidades ***
						selecao = " * "
						tabela = " TBL_unidades"
						condicoes = " WHERE cd_status > 0 "'*** Status zero é inativo ***
						criterios = " Order by nm_unidade"

						strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
										  	
						Set rs_uni = dbconn.execute(strsql)%>
					<td bgcolor="#f5f5f5" valign="top" colspan=4>
					<select name="nm_unidade" class="inputs" onFocus="nextfield ='num_qtd';">
					<option value=""></option>
						<%
						Do while not rs_uni.EOF
						cd_uni = rs_uni("cd_codigo")
						nm_sigla = rs_uni("nm_sigla")
						nome_unidade = rs_uni("nm_unidade")
						
						if nm_unidade = nm_sigla Then
						ck_uni = "selected"
						end if%>
						
					<option value="<%=nm_sigla%>" <%=ck_uni%>><%=nm_sigla%> - <%=nome_unidade%></option>
						<%rs_uni.movenext
						ck_uni = ""
						'salto = salto + 1
						loop
						%>
					</select>
			

					</td>
				</tr>
				
			<!--Fim da etapa 1.1-->			
			<!--Início da etapa 1.2-->				
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>
				
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>	
				</tr>
				<!--Os novos campos começam aqui-->
				<%if esconder = "1" Then
				Else%>
				<%end if%>
				<tr>
					<td colspan="100%" align="center">
					<br><%if acao = "1" Then%><input  name="ok" type="submit" value="Cadastrar O.S."><%else%>
					<input type="submit" value="Modificar O.S.">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ordemservico_andamento.asp?cd_codigo=<%=cd_codigo%>&campo=<%=campo%>">Voltar</a><%End if%>
					
					</td></tr>
				<!--Fim da etapa 1.2-->
			
						</form>
							
			</table>

		</td>
	</tr>
</table>	
</BODY>
</HTML>



