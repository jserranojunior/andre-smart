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

<%'cd_unidade = request("cd_unidade")
'cd_protocolo = request("cd_protocolo")
	'protocolo = request("protocolo")
'cd_digito = request("cd_digito")

pacientes_dvd = request("pacientes_dvd")
response.write(pacientes_dvd)

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

'if action_form = "" Then
'	action_form = "protocolo.asp"
'elseif action_form <> "" Then
'	action_form = "acoes/protocolo_acao.asp"
'end if
action_form = "protocolo.asp"

'modo = "ins"
%>
<!--#include file="../js/ajax.js"-->
<!--#include file="../js/protocolo.js"-->

<BODY onLoad="foco(cd_unidade)">
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
					/ <a href="protocolo.asp?tipo=digitacao">Alterar</a> 
					/ <a href="protocolo.asp?tipo=ver">Consultar</a>
			</td>
		</tr>
		<tr id="no_print"><td colspan="100%"><img src="imagens/px.gif" alt="" width="100%" height="5" border="0"></td></tr>	
		<tr bgcolor="#808080">
			<td align="center" colspan="100%">&nbsp;<font size="2" color="white"><b>Dados das cirurgias</b></font></td>
		</tr>
		<form method="post" action="<%=action_form%>" name="Form" id="form">
		<input type="hidden" name="modo" value="<%=modo%>">
		<input type="hidden" name="tipo" value="<%=tipo%>">
		<input type="hidden" name="action_form" value="ok">
		<!--input type="hidden" name="acao" value="<%'=acao%>"-->


<%'if pacientes_dvd = "" Then
'*********************************************************************
'***** 1.0 envia dados dos pacientes para avaliação 			 *****
'*********************************************************************%>
		<tr bgcolor="#f5f5f5">
			<td align="left" colspan="1">&nbsp;Pacientes  <font color="red"><b><%=alerta_unid%></b></font></td>
			<td colspan="3">&nbsp;
				<input type="hidden" name="acao" value="<%=acao%>">
				<input type="text" name="pacientes_dvd" size="25" maxlength="50" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 50)" onFocus="PararTAB(this);" class="inputs" value="<%=protocolo%>">
				<!--input type="text" name="cd_digito" size="1" maxlength="1" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 1)" onFocus="PararTAB(this);" onBlur="document.form.submit()" class="inputs" value="<%=cd_digito%>">&nbsp;-->				
				<br><%=mensagem%>
			</td>
		</tr>
		<tr>
		<td colspan="100%" align="center" valign="center">
		<%'if cd_prot = "0" OR cd_prot = "" Then%>
			<input type="submit" name="Submit" value="ok"> &nbsp; 
		<%'end if%>
		<a href="protocolo.asp?tipo=<%=tipo%>">Limpar</a></td>
	</tr>
	
			</form>

<%if pacientes_dvd <> "" Then
'**********************************************
'***** 1.2 Mostra os dados das cirurgias  *****
'**********************************************
			'looping para cada nº de paciente
			numero = 1
			proc_array = split(pacientes_dvd,",")
			For Each proc_item In proc_array
			
				a002_pacreg = proc_item'pacientes_dvd
			
			xsql ="SELECT * FROM TBL_protocolo WHERE A002_pacreg = '"&a002_pacreg&"'"
					Set rs = dbconn.execute(xsql)
					do while Not rs.EOF
					cd_codigo = rs("A002_numseq")
					'protocolo = rs("A002_numpro")
					'cd_prot = rs("A002_numpro")
					'cd_unidade = rs("A053_codung")
					nm_nome = rs("A002_pacnom")
					cd_registro = rs("A002_pacreg")
					'nm_cirug_realizada = rs("A002_limarm")
					dt_procedimento = rs("A002_datpro")
					'dt_hr_agenda = rs("A002_horage")
					dt_inicio = rs("A002_horini")
						hora_inicio = hour(dt_inicio)&minute(dt_inicio)						 
					dt_fim = rs("A002_horfin")
						hora_fim = hour(dt_fim)&minute(dt_fim)					
							if hora_inicio <> 0 and hora_fim <> 0 then
								dt_duracao = total_mes(datediff("n",dt_inicio,dt_fim))'&" ["&dt_inicio&" - "&dt_fim&"]"
								negativo = Instr(1,dt_duracao,"-",0)
								if negativo > 0 Then
								dt_duracao = replace(dt_duracao,"-","")
								end if
							end if
					
					cd_med = rs("A055_codmed")
					'cd_especialidade = rs("A057_codesp")
					rs.movenext
					loop
						'if cd_codigo <> "" Then%>
							
		<tr bgcolor="#f5f5f5">
			<td colspan="100%">&nbsp;</td>
		</tr>
		<!--tr bgcolor="#808080">
			<td align="center" colspan="100%"><font size="2" color="white"><b>&nbsp;Identificação do Paciente</b></font></td>
		</tr-->	
	<%if cd_registro <> "" Then%>
		<tr bgcolor="#e0e0e0">
			<td rowspan="4" align="center" valign="center"><%=numero%></td>
			<td align="left" valign="top" colspan="1" width="100"><b>Paciente</b></td>
			<td align="left" valign="top" colspan="1" width="80"><%=cd_registro%></td>
			<td align="left" valign="top" colspan="1" width="300"><%=nm_nome%></td>
		</tr>
		
			<%strsql ="Select * from TBL_medicos where a055_codmed = '"&cd_med&"'"
			Set rs_med = dbconn.execute(strsql)
			Do while Not rs_med.EOF%>
		<tr bgcolor="#e0e0e0">
			<td align="left" valign="top" colspan="1"><b>Médico</b></td>
			<td align="left" valign="top" colspan="1"><%'=cd_med%><%=rs_med("a055_crmmed")%></td>
			<td align="left" valign="top" colspan="1"><%=rs_med("a055_desmed")%></td>
		</tr>
			<%rs_med.movenext
			Loop%>
		<tr bgcolor="#e0e0e0">
			<td align="left" valign="top" colspan="1"><b>Data da cirurgia</b></td>
			<td align="left" valign="top" colspan="1"><%=dt_procedimento%></td>
			<td align="left" valign="top" colspan="1"><%'=hour(dt_duracao)%> <%'=minute(dt_duracao)%> <%'=dt_inicio &" "&dt_fim%></td>
		</tr>
		<tr bgcolor="#e0e0e0" bgcolor="#e0e0e0">
			<td align="left" valign="top" valign="top"><b>Procedimentos</b></td>
			<td align="left" valign="top" colspan="2">
					<table border="0">
					<%'*** Procedimentos ***
					xsql_proc ="SELECT * FROM TBL_protocolo_procedimento WHERE A002_numseq = '"&cd_codigo&"' order by a003_propri"
					Set rs_proc = dbconn.execute(xsql_proc)
					while Not rs_proc.EOF
					cd_procedimento = rs_proc("A058_codpro")%>
							
						<%cd_procedimento = replace(cd_procedimento,","," OR a058_codpro=")%>
						<%strsql ="Select * from TBL_procedimento where a058_codpro="&cd_procedimento&""
						Set rs = dbconn.execute(strsql)%>		
						<%Do while not rs.EOF%>	
						<tr>
							<td align="left" valign="top" bgcolor="#e0e0e0" width="78"><%=cd_procedimento%></td>
							<td align="left" valign="top" bgcolor="#e0e0e0" width="300"><%=rs("a058_despro")%></td>
						</tr>
						<%rs.movenext
						loop%>
					<%rs_proc.movenext
					wend%>
						</table>
			</td>
		</tr>
	<%Else%>
		<tr bgcolor="#e0e0e0">
			<td align="center" valign="center"><%=numero%></td>
			<td align="left" valign="top" colspan="3" width="300"><b>O registro <%=a002_pacreg%> ainda não está disponível.</b></td>
		</tr>
	<%end if%>
		<%numero = numero + 1
		Next
		
		response.write("<tr><td colspan=4><br><br>&quot;Fim da Consulta&quot;</td></tr>")%>
		
	</table>
</td>
</tr>
					
					
					
							<!--tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="6" height="1" border="0"></td></tr-->
					<%'end if%>
<%



'**************************************
'***** Dados do protocolo - 2ª fase *****

'////////////////////////////////////		
end if%>
	<br>
<br>
</BODY>
</HTML>


