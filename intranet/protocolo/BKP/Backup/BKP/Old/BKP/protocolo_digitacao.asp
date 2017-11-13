<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<script language="Javascript">
<!-- 
VerifiqueTAB=true;
function Mostra(quem, tammax) {
if ( (quem.value.length == tammax) && (VerifiqueTAB) ) { 
var i=0,j=0, indice=-1;
for (i=0; i<document.forms.length; i++) { 
for (j=0; j<document.forms[i].elements.length; j++) { 
if (document.forms[i].elements[j].name == quem.name) { 
indice=i;
break;
} 
} 
if (indice != -1) break; 
} 
for (i=0; i<=document.forms[indice].elements.length; i++) { 
if (document.forms[indice].elements[i].name == quem.name) { 
while ( (document.forms[indice].elements[(i+1)].type == "hidden") &&
(i < document.forms[indice].elements.length) ) { 
i++;
} 
document.forms[indice].elements[(i+1)].focus();
VerifiqueTAB=false;
break;
} 
} 
} 
} 

function PararTAB(quem) { VerifiqueTAB=false; } 
function ChecarTAB() { VerifiqueTAB=true; } 


function foco(){
document.forms[1].elements[0].focus();
}

//---------------------------------------------------
function soma(a,b,c,procedimentos){
	a=a;
	b=b;
	c=c;
	procedimentos=procedimentos
	
		if (procedimentos != ''){
			virg =  ',';
			//virg = ',\r\n';
			}
		else{
		virg= ''
		}		
	procedimentos = procedimentos + virg
		if (a != ""){
			procedimentos = procedimentos + a;
			}
		else{
			procedimentos = procedimentos + b;
			}
		if (c != ""){
			//materiais = materiais - c;
			procedimentos = procedimentos.replace(c,'');
			}	
		
	document.form.res.value=a;
	document.form.procedimentos.value=(procedimentos.replace(",,", ","));
	document.form.cd_procedimento.value="";
	document.form.cd_procedimento_1.value="";
	document.form.cd_procedimento_2.value="";
		
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
//-------------------------------------------------------
function soma_mat(j,k,l,materiais){
	j=j;
	k=k;
	l=l;
	materiais=materiais
	
		if (materiais != ''){
			virg_m =  ',';			
			}
		else
		{
		virg_m= ''
		}		
	materiais = materiais + virg_m
		if (j != ""){
			materiais = materiais + j;
			}
		else 
			{
			materiais = materiais + k;
			}
		if (l != ""){
			materiais = materiais.replace(l,'');
			}

	document.Form.res_mat.value=j;
	document.Form.materiais.value=(materiais.replace(",,", ","));
	document.Form.cd_material.value="";
	document.Form.cd_material_1.value="";
	document.Form.cd_material_2.value="";	
{
	var oHTTPRequest_mat = createXMLHTTP(); 
    oHTTPRequest_mat.open("post", "../ajax/material/ajax_materiais.asp", true);
    oHTTPRequest_mat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    oHTTPRequest_mat.onreadystatechange=function(){
     if (oHTTPRequest_mat.readyState==4){
        document.all.divMatc.innerHTML = oHTTPRequest_mat.responseText;}}
       oHTTPRequest_mat.send("materiais=" + form.materiais.value);
 }
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
	'protocolo = request("protocolo")
cd_digito = request("cd_digito")

cd_codigo = request("cd_codigo")
tipo = request("tipo")
modo = request("modo")
mensagem = request("mensagem")

'if cd_codigo = "" AND modo = "ins" Then
if modo = "ins" Then
acao = "inserir"
'Elseif cd_codigo <> "anything" AND modo = "ver" then
Elseif modo = "ver" then
acao = "editar"
end if

action_form = request("action_form")

if action_form = "" Then
	action_form = "protocolo.asp"
elseif action_form <> "" Then
	action_form = "acoes/protocolo_acao.asp"
end if
'modo = "ins"
%>
<!--#include file="../js/ajax.js"-->
<!--#include file="../js/protocolo.js"-->

<BODY onLoad="foco(Form.cd_unidade)">
<!--body--><br>
<!--Cabeçalho do formulário -->
	<table  border="1" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno" align="center">
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
		<form method="post" action="<%=action_form%>" name="Form" id="form">
		<input type="hidden" name="modo" value="<%=modo%>">
		<input type="hidden" name="tipo" value="<%=tipo%>">
		<input type="hidden" name="action_form" value="ok">
		<!--input type="hidden" name="acao" value="<%'=acao%>"-->


<%if cd_unidade = "" AND cd_protocolo = "" Then
'********************************************************************
'***** 1.0 envia dados do protocolo para avaliação (ver/editar) *****
'********************************************************************%>
		<tr bgcolor="#f5f5f5">
			<td align="left">&nbsp;Protocolo  <font color="red"><b><%=alerta_unid%></b></font></td>
			<td>&nbsp;
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
					do while Not rs.EOF
					unidade = rs("cd_codigo")
					rs.movenext
					loop
						if unidade = "" Then 'Se a unidade não existir, retorna aos campos iniciais
							response.redirect("protocolo.asp?tipo=ver&mensagem=(Unidade Inválida)&protocolo="&cd_protocolo&"")
						end if
						
						
'*********************************************
'***** 1.2 Verifica o digito verificador *****
'*********************************************
				'verifica_digito = digitov(zero(int(cd_unidade))&"."&proto(int(cd_protocolo)))
				'response.write(digitov("03.114147")&"<br>")
				verifica_digito = digitov(cd_unidade&"."&proto(cd_protocolo))
				
				if right(verifica_digito,1) = cd_digito then
					response.write("número correto: "&verifica_digito)
				else
				response.write("Numero inválido <br>")
				response.write("número correto: "&verifica_digito)
				response.redirect("protocolo.asp?tipo="&tipo&"&mensagem=(Número Inválido)")
				end if
				
'*******************************************************
'***** 1.3 Verifica se o nº do protocolo já existe *****
'*******************************************************

			xsql ="SELECT * FROM TBL_protocolo WHERE A002_numpro = '"&cd_protocolo&"' AND A053_codung = '"&cd_unidade&"'"
					Set rs = dbconn.execute(xsql)
					do while Not rs.EOF
					
						cd_codigo = rs("a002_numseq")
						protocolo = rs("A002_numpro")
						cd_prot = rs("A002_numpro")
						cd_unidade = rs("A053_codung")
						nm_nome = rs("A002_pacnom")
						cd_idade = rs("A002_pacida")
						cd_registro = rs("A002_pacreg")
						cd_convenio = rs("A054_codcon")
						nm_sexo = rs("A002_pacsex")
						nm_cirug_realizada = rs("A002_limarm")
						dt_procedimento = rs("A002_datpro")
						nm_agenda = rs("A002_tipage")
						dt_hr_agenda = rs("A002_horage")
						dt_inicio = rs("A002_horini")
						dt_fim = rs("A002_horfin")
							dt_duracao = total_mes(datediff("n",dt_inicio,dt_fim))'&" ["&dt_inicio&" - "&dt_fim&"]"
						cd_med = rs("A055_codmed")
						cd_rack = rs("a056_codrac")
						cd_especialidade = rs("A057_codesp")
						'cd_procedimento = rs("cd_procedimento")
					
					rs.movenext
					loop
						'if cd_codigo <> "" Then%>
							
							<%if tipo = "digitacao" AND cd_codigo <> "0" then%>
							
							<%some = 0%>
							<!--#include file="protocolo_dig.asp"-->
							
							<%elseif tipo = "ver" AND cd_codigo <> "" then%>
							<%some = 0%>
							<!--#include file="protocolo_ver.asp"-->
							
							<%elseif tipo = "dvd" AND cd_codigo <> "" then%>
							<%some = 0%>
							<!--#include file="protocolo_dvd.asp"-->
							
							<%elseif tipo = "ver" AND cd_codigo = "" then%>
							<%some = 0%>
							<!--include file="protocolo_ver.asp"-->
							<tr><td align="center"><br>**** O protocolo solicitado ainda não foi digitado.***<br>
									Tente alterar os dados da pesquisa<br>&nbsp;									
								</td></tr>
							<%end if%>
							<!--tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="6" height="1" border="0"></td></tr-->
					<%'end if%>
<%



'**************************************
'***** Dados do protocolo - 2ª fase *****

'////////////////////////////////////		
end if%>
	<tr>
		<td colspan="100%" align="center" valign="center">
		<%if some <> "1" Then%>
			<input type="submit" name="Submit" value="ok"> &nbsp; 
		<%end if%>
		<a href="protocolo.asp?tipo=<%=tipo%>">Limpar</a></td>
	</tr>
	<tr><td style="width:260px;">&nbsp;.</td><td style="width:280px;" colspan="2">&nbsp;.</td><td style="width:150px;">&nbsp;.</td></tr>
			</form>
	<!--tr><td><img src="../imagens/blackdot.gif" alt="" width="1" height="1" border="0"></td></tr-->
</table><br>
<br>
</BODY>
</HTML>


