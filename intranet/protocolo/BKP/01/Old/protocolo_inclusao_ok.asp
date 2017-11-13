<!--#include file="../includes/util.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<script language="Javascript">
<!-- 
function foco(){
document.forms[1].elements[0].focus();
}

function soma(a,b,procedimentos){
	a=a;
	b=b;
	
		if (procedimentos != ''){
			virg =  ',';
			//virg = ',\r\n';
			}
		else
		{
		virg= ''
		}
		
	procedimentos = procedimentos + virg
		if (a != ""){
			procedimentos = procedimentos + a;
			}
		else
			{
			procedimentos = procedimentos + b;
			}

	document.form.res.value=a;
	//document.form.procedimentos.value=procedimentos;
	document.form.procedimentos.value=(procedimentos.replace(",,", ","));
	document.form.cd_procedimento.value="";
	document.form.cd_procedimento_1.value="";

{
     var oHTTPRequest_prot = createXMLHTTP(); 
     oHTTPRequest_prot.open("post", "../ajax/procedimento/ajax_procedimentos.asp", true);
     oHTTPRequest_prot.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_prot.onreadystatechange=function(){
      if (oHTTPRequest_prot.readyState==4){
         document.all.divprot.innerHTML = oHTTPRequest_prot.responseText;}}
       oHTTPRequest_prot.send("procedimentos=" + form.procedimentos.value);
 }


}


function mostra_convenio_()
 {
     var oHTTPRequest_conv = createXMLHTTP(); 
     oHTTPRequest_conv.open("post", "../ajax/protocolo/ajax_convenio.asp", true);
     oHTTPRequest_conv.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_conv.onreadystatechange=function(){
      if (oHTTPRequest_conv.readyState==4){
         document.all.divAssist_m1.innerHTML = oHTTPRequest_conv.responseText;}}
       oHTTPRequest_conv.send("cd_convenio=" + Form_proto.cd_convenio.value);
 }
// -->	

<!-- Begin
nextfield = "nm_nome"; // nome do primeiro campo do site
//netscape = "";
ver = navigator.appVersion; 
len = ver.length;
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
			document.write(k);
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

<%cd_unidade = request("cd_unidade")
cd_protocolo = request("cd_protocolo")
cd_digito = request("cd_digito")

cd_codigo = request("cd_codigo")
tipo = request("tipo")
modo = request("modo")
mensagem = request("mensagem")

if modo = "ins" Then
acao = "inserir"
Elseif modo = "ver" then
acao = "editar"
end if

action_form = request("action_form")

'if action_form = "" Then
	action_form = "protocolo.asp"
'elseif action_form <> "" Then
'	action_form = "acoes/protocolo_acao.asp"
'end if

%>
<!--#include file="../js/ajax.js"-->
<!--#include file="../js/protocolo.js"-->

<BODY>
<!--body--><br>
<!--Cabeçalho do formulário -->
	<table width="500" border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno" align="center">
		<tr>		
			<td class="txt_cinza" colspan="100%">
			<b>Protocolos &raquo; <%=tipo%><br><br></td>
			</td>
		</tr>
		<tr>
			<td colspan="100%"><!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					/ <a href="protocolo.asp?tipo=digitacao">Novo</a>
					/ <a href="protocolo.asp?tipo=digitacao&cd_protocolo=<%=cd_protocolo%>&cd_unidade=<%=cd_unidade%>">Alterar</a> 
					/ <a href="protocolo.asp?tipo=ver">Consultar</a>
			</td>
		</tr>
		<tr id="no_print"><td colspan="100%"><img src="imagens/px.gif" alt="" width="100%" height="5" border="0"></td></tr>	
		<tr bgcolor="#808080">
			<td align="center" colspan="100%">&nbsp;<font size="2" color="white"><b>Dados do Protocolo</b></font></td>
		</tr>
		<form method="post" action="<%=action_form%>" name="Form_proto" id="form">
		<input type="hidden" name="modo" value="<%=modo%>">
		
		<input type="hidden" name="action_form" value="ok">
		<!--input type="hidden" name="acao" value="<%'=acao%>"-->


<%if cd_unidade = "" AND cd_protocolo = "" Then
'********************************************************************
'***** 1.0 envia dados do protocolo para avaliação (ver/editar) *****
'********************************************************************%>
		<tr bgcolor="#f5f5f5">
			<td align="left">&nbsp;Protocolo  <font color="red"><b><%=alerta_unid%></b></font></td>
			<td>&nbsp;
				<input type="hidden" name="tipo" value="inclusao">
				<input type="hidden" name="acao" value="<%=acao%>">
				<input type="text" name="cd_unidade" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" class="inputs" value="<%=cd_unidade%>">.
				<input type="text" name="cd_protocolo" size="6" maxlength="6" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 6)" onFocus="PararTAB(this);" class="inputs" value="<%=protocolo%>">
				<input type="text" name="cd_digito" size="1" maxlength="1" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 1)" onFocus="PararTAB(this);" onBlur="document.form.submit()" class="inputs" value="<%=cd_digito%>">				
				<br><%=mensagem%>
			</td>
		</tr>
<%elseif cd_unidade <> "" AND cd_protocolo <> "" Then
'**********************************************
'***** 1.1 Verifica se a unidade é válida *****
'**********************************************
			'*** 1.a - Verifica a unidade ***
			xsql ="SELECT * FROM TBL_unidades WHERE cd_codigo = '"&cd_unidade&"'"
					Set rs = dbconn.execute(xsql)
						if rs.eof Then 'Se a unidade não existir, retorna aos campos iniciais
							response.redirect("protocolo.asp?tipo=inclusao&mensagem=(Unidade Inválida)&protocolo="&cd_protocolo&"")
							'response.write("Unidade inválida")
						end if
						
						while Not rs.EOF
						cd_unidade = rs("cd_codigo")
						rs.movenext
						wend
						
'**********************************************************************************************
'***** 1.3 Verifica se o nº do protocolo é já foi impresso e se já existe na base de dados*****
'**********************************************************************************************
			'*** 1.a - Verifica se o protocolo já foi digitado ***
			'xsql ="SELECT * FROM TBL_protocolo_etiquetas WHERE cd_unidade = cd_unidadea053_codung = '"&cd_unidade&"'"
			'		Set rs = dbconn.execute(xsql)
			'		do while Not rs.EOF
			'		unidade = rs("a053_codung")
			'		rs.movenext
			'		loop
			'			if unidade = "" Then 'Se a unidade não existir, retorna aos campos iniciais
			'				'response.redirect("protocolo.asp?tipo=ver&mensagem=(Unidade Inválida)&protocolo="&cd_protocolo&"")
			'			end if

			xsql ="SELECT * FROM TBL_protocolo WHERE cd_protocolo = '"&cd_protocolo&"' AND cd_unidade = "&cd_unidade&""
					Set rs = dbconn.execute(xsql)
						if rs.EOF Then 'Se o numero for invalido, retorna aos campos iniciais
							'response.write("Numero inválido:"&cd_unidade&"."&cd_protocolo)
							response.write("Numero ainda não foi digitado.<br>")
						else
							'response.redirect("protocolo.asp?tipo=inclusao&mensagem=(Numero Invalido)&cd_unidade="&cd_unidade&"")
							response.write("Numero Já existe: "&zero(cd_unidade)&"."&proto(cd_protocolo))
						end if
						
						'while Not rs.EOF
						'	cd_protocolo = rs("cd_protocolo")
						'rs.movenext
						'wend
						'	if unidade = "" Then 'Se a unidade não existir, retorna aos campos iniciais
						'		'response.redirect("protocolo.asp?tipo=ver&mensagem=(Unidade Inválida)&protocolo="&cd_protocolo&"")
						'	end if


'*********************************************
'***** 1.2 Verifica o digito verificador *****
'*********************************************
				'digito_correto = digitov(zero(cd_unidade)&proto(cd_protocolo))
				
				digito_verifica = digitov(zero(cd_unidade)&proto(cd_protocolo))
				digito_correto = mid(digito_verifica,10,1)
				
				
				if int(digito_correto) = int(cd_digito) then
					
					response.write("número correto!<br>::"&digito_verifica&"::")
					'response.redirect("protocolo.asp?tipo=formulario&num_protocolo="&zero(cd_unidade)&%>					
										
				<%else
				response.write("<br>Digito correto :"&digito_correto&"<br>")
				response.write("OK :"&zero(cd_unidade)&"."&cd_protocolo&"-"&cd_digito)
				end if
				'response.write("<br>Digito :"&digito_correto&"<br>")


end if%>
	<tr>
		<td colspan="100%" align="center" valign="center">
		<%'if cd_prot = "0" OR cd_prot = "" Then%>
			<input type="submit" name="Submit" value="ok"> &nbsp; 
		<%'end if%>
		<a href="protocolo.asp?tipo=<%=tipo%>">Limpar</a></td>
	</tr>
	<!--#include file="protocolo_dig.asp"-->
	
			</form>
<!-- aqui é o grande segredo. Essa função -->

</table><br>

<br>
</BODY>
</HTML>


