<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<%'Response.AddHeader "Content-Type", "text/html; charset=iso-8859-1"%>
<%Response.Charset="ISO-8859-1" %>
 
<html>
<head>

</head>
<!--script language="JavaScript" type="text/javascript" src="js/richtext.js"></script-->
<!--script language="JavaScript" type="text/javascript" src="../js/mascarainput.js"></script-->

<script language="JavaScript">
function validar_protocolo(shipinfo){
	//alert(shipinfo);
	//formulario = document.getElementById("form").name
	/*hoje = new Date
	dia_hoje = hoje.getDate();
		if (dia_hoje < 10 ) dia_hoje = '0'+ dia_hoje
	mes_hoje = hoje.getMonth()+1;
		if (mes_hoje < 10 ) mes_hoje = '0'+ mes_hoje
	ano_hoje = hoje.getFullYear();
		data_hoje = ano_hoje +''+ (mes_hoje) +''+ dia_hoje;
	//********************************************************	
	data_informada = shipinfo.dt_ano_cirurgia.value +''+ shipinfo.dt_mes_cirurgia.value +''+ shipinfo.dt_dia_cirurgia.value;
	dif_datas = data_hoje-data_informada;
	//alert(data_informada+' ~ '+data_hoje+'='+(data_hoje-data_informada));
	
	
	if (shipinfo.cd_libera_data.value=0) { alert("Liberação");shipinfo.dt_mes_cirurgia.focus(shipinfo.dt_mes_cirurgia.value="");shipinfo.dt_dia_cirurgia.focus(shipinfo.dt_dia_cirurgia.value="");return (false);}
	
	if (data_informada>data_hoje){window.alert ("A data preenchida está incorreta.");shipinfo.dt_ano_cirurgia.value="";shipinfo.dt_mes_cirurgia.value="";shipinfo.dt_dia_cirurgia.value="";shipinfo.dt_dia_cirurgia.focus(); return (false);}
	if (dif_datas > 101)&&(shipinfo.cd_libera_data.value=0) { alert("A data informada é muito antiga");shipinfo.dt_mes_cirurgia.focus(shipinfo.dt_mes_cirurgia.value="");shipinfo.dt_dia_cirurgia.focus(shipinfo.dt_dia_cirurgia.value="");return (false);}
	*/
//alert(formulario)
	//if (shipinfo.nm_nome.value.length==""){window.alert ("O Nome deve ser informado.");shipinfo.nm_nome.focus();return (false);}	
	//if (shipinfo.cd_idade.value.length==""){window.alert ("A idade deve ser informada.");shipinfo.cd_idade.focus();return (false);}
	/*if (shipinfo.nm_registro.value.length==""){window.alert ("O número do Paciente deve ser informado.");shipinfo.nm_registro.focus();return (false);}	
	if (shipinfo.cd_convenio_x.value.length==""){window.alert ("O convenio deve ser informado.");shipinfo.cd_convenio.focus();return (false);}	
	if (shipinfo.nm_sexo.value.length==""){window.alert ("O sexo deve ser informado.");shipinfo.nm_sexo.focus();return (false);}	
	if (shipinfo.nm_cirug_realizada.value.length==""){window.alert ("Informe se a cirurgia foi realizada.");shipinfo.nm_cirug_realizada.focus();return (false);}	
	if (shipinfo.dt_dia_cirurgia.value.length==""){window.alert ("A data deve ser informada.");shipinfo.dt_dia_cirurgia.focus();return (false);}
	if (shipinfo.dt_mes_cirurgia.value.length==""){window.alert ("A data deve ser informada.");shipinfo.dt_mes_cirurgia.focus();return (false);}
	if (shipinfo.dt_ano_cirurgia.value.length==""){window.alert ("A data deve ser informada.");shipinfo.dt_ano_cirurgia.focus();return (false);}
	if (shipinfo.dt_dia_cirurgia.value>=32){window.alert ("O dia da cirurgia está inválido.");shipinfo.dt_dia_cirurgia.focus(shipinfo.dt_dia_cirurgia.value="");return (false);}
	if (shipinfo.dt_mes_cirurgia.value>=13){window.alert ("O mês da cirurgia está inválido.");shipinfo.dt_mes_cirurgia.focus(shipinfo.dt_mes_cirurgia.value="");return (false);}
	if (shipinfo.dt_ano_cirurgia.value<2000){window.alert ("O ano da cirurgia está inválido.");shipinfo.dt_ano_cirurgia.focus(shipinfo.dt_ano_cirurgia.value="");return (false);}
	
	//if (data_informada > data_hoje){ alert("A data informada é maior que a data atual");shipinfo.dt_mes_cirurgia.focus(shipinfo.dt_mes_cirurgia.value="");shipinfo.dt_dia_cirurgia.focus(shipinfo.dt_dia_cirurgia.value="");return (false);}
	if (shipinfo.nm_agenda_cirurgia.value<2000){window.alert ("O tipo de agenda deve ser informado.");shipinfo.nm_agenda_cirurgia.focus(shipinfo.nm_agenda_cirurgia.value="");return (false);}
	if (shipinfo.dt_hora_inicio.value.length==""){window.alert ("O horário do início da cirurgia deve ser informado.");shipinfo.dt_hora_inicio.focus();return (false);}
	if (shipinfo.dt_minuto_inicio.value.length==""){window.alert ("O horário do início da cirurgia deve ser informado.");shipinfo.dt_minuto_inicio.focus();return (false);}
	if (shipinfo.dt_hora_inicio.value>=24){window.alert ("O horário do início da cirurgia está inválido.");shipinfo.dt_hora_inicio.focus(shipinfo.dt_hora_inicio.value="");return (false);}
	if (shipinfo.dt_minuto_inicio.value>=60){window.alert ("O horário do início da cirurgia está inválido.");shipinfo.dt_minuto_inicio.focus(shipinfo.dt_minuto_inicio.value="");return (false);}
	if (shipinfo.dt_hora_fim.value.length==""){window.alert ("O horário do término da cirurgia deve ser informado.");shipinfo.dt_hora_fim.focus();return (false);}
	if (shipinfo.dt_minuto_fim.value.length==""){window.alert ("O horário do término da cirurgia deve ser informado.");shipinfo.dt_minuto_fim.focus();return (false);}
	if (shipinfo.dt_hora_fim.value>=24){window.alert ("O horário do término da cirurgia está inválido.");shipinfo.dt_hora_fim.focus(shipinfo.dt_hora_fim.value="");return (false);}
	if (shipinfo.dt_minuto_fim.value>=60){window.alert ("O horário do término da cirurgia está inválido.");shipinfo.dt_minuto_fim.focus(shipinfo.dt_minuto_fim.value="");return (false);}
	if (shipinfo.cd_crm_x.value.length==""){window.alert ("O medico deve ser informado.");shipinfo.cd_crm.focus();return (false);}	
	if (shipinfo.cd_rack_1.value.length==""){window.alert ("O rack deve ser informado.");shipinfo.cd_rack.focus();return (false);}	
	if (shipinfo.cd_especialidade_x.value.length==""){window.alert ("A especialidade deve ser informada.");shipinfo.cd_especialidade.focus();return (false);}	
	if (shipinfo.procedimentos_total.value.length==""){window.alert ("Um procedimento deve ser informado.");shipinfo.cd_procedimento_busca.focus();return (false);}	
	*/
	return (true);
}	
	
</script>

<!--script language="Javascript">

</script-->

<script language="Javascript">
<!-- 
	
	
function detecta_enter(event,prox_campo,usuario) {
    var char = event.which || event.keyCode;
	//var char = event.keyCode;
	
    if (char == 13 && prox_campo != "submit" && usuario != "4600"){
		//alert(char + prox_campo);
		//alert(usuario + "_" + prox_campo)
		document.getElementById(prox_campo).focus();
		return (true); 
	}
	if (char == 13 && prox_campo == "submit" && usuario != "4600"){
		//alert(usuario + "_" + prox_campo)
		document.getElementById("form").submit();
		return (true);
	}
	
}



/*VerifiqueTAB=true;
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
*/
/*
function foco(){
document.forms[1].elements[0].focus();
}
*/
//---------------------------------------------------

function alimenta_convenio(cod){
	document.getElementById("cd_convenio_x").value=cod;
	}
function alimenta_medico(cod){
	document.getElementById("cd_crm_x").value=cod;
	}
function alimenta_rack(cod){
	document.getElementById("cd_rack_x").value=cod;
	}
function alimenta_especialidade(cod){
	document.getElementById("cd_especialidade_x").value=cod;
	}

function foca_unidade(){
		document.getElementById("cd_unidade").focus();
	}
function foca_nome(){
		document.getElementById("nm_nome").focus();
	}		

function verifica_data(){
	//if (document.Form.dt_ano_cirurgia.lenth < 2) alert (diferenca_datas);
	
	hoje = new Date
		dia_hoje = hoje.getDate();
			if (dia_hoje < 10 ) dia_hoje = '0'+ dia_hoje
		mes_hoje = hoje.getMonth()+1;
			if (mes_hoje < 10 ) mes_hoje = '0'+ mes_hoje
		ano_hoje = hoje.getFullYear();
			data_hoje = ano_hoje +''+ (mes_hoje) +''+ dia_hoje;	
	//data_hoje = ano_hoje+mes_hoje+dia_hoje
	data_hoje = ano_hoje+mes_hoje+'01'
	data_cirurg = document.Form.dt_ano_cirurgia.value+document.Form.dt_mes_cirurgia.value+document.Form.dt_dia_cirurgia.value;
	
	diferenca_datas = data_hoje-data_cirurg
	
	
	//if (document.Form.dt_mes_cirurgia.value.length > 1) {	
	//if (diferenca_datas > 70 && document.Form.dt_mes_cirurgia.value.length > 1) {	
		//alert(data_cirurg+'/'+data_hoje+'='+diferenca_datas);}
		//alert("A data provavelmente está incorreta");
		//document.Form.dt_mes_cirurgia.value="";
		//}
	}
	
//---------------------------------------------------

function soma(proc1,proc_apagar,procedimentos_total){

		if (procedimentos_total != ''){
			virg =  ',';
			}
		else{
			virg= ''
			}
	
	procedimentos_total = procedimentos_total + virg
		if (proc1 != ""){
			//procedimentos_total = procedimentos_total + proc1;
			procedimentos_total = procedimentos_total + proc1;
			}
		if (proc_apagar != ""){
			procedimentos_total = procedimentos_total.replace(proc_apagar,'');
			}
			
		
	document.Form.proc_res.value=proc1;
	document.Form.procedimentos_total.value=(procedimentos_total.replace(",,", ","));
	document.Form.cd_procedimento_busca.value="";
	//document.Form.cd_procedimento_busca.focus();
	//document.Form.cd_procedimento_1.value="";
	//document.Form.cd_procedimento_2.value="";
	
	//chama Ajax	
	loadXMLDoc("../ajax/procedimento/ajax_procedimentos_lista.asp?procedimentos_total="+procedimentos_total,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
						document.getElementById("divproc_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		});	

}
//-------------------------------------------------------
function soma_mats(){
	alert("teste");
}

function soma_mat(material,quantidade,mat_apagar,materiais_total){
	if (materiais_total != ''){
		virg_m =  ',';			
		}
	else
		{
		virg_m= ''
		}		
	materiais_total = materiais_total + virg_m
		if (material != ""){
			materiais_total = materiais_total + material + ';' + quantidade;
			}
		if (mat_apagar != ""){
		//alert(mat_apagar+';'+quantidade);
			mat_aagar = mat_apagar+';'+quantidade
			materiais_total = materiais_total.replace(mat_apagar+';'+quantidade,'');
			material = ''
			cd_material_busca = ''
			}

	document.Form.res_mat.value=material;
	document.Form.materiais_total.value=(materiais_total.replace(",,", ","));
	document.Form.cd_material_busca.value="";
	document.Form.qtd_material.value="";
	//document.Form.cd_material_1.value="";
	//document.Form.cd_material_2.value="";
	
//{
	
	loadXMLDoc("../ajax/material/ajax_materiais_lista.asp?materiais_total="+materiais_total,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
						
						document.getElementById("divMat_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		});
 //}
}

//***********************************************************************************
function soma_patrimonios(){
	alert("teste")
	}

function soma_patrimonio(patrimonio,pat_apaga,subtotal_patrimonio){
	if (subtotal_patrimonio != ''){
		virg_pat = ',';
		}
	else
		{
		virg_pat= ''
		}
		subtotal_patrimonio=subtotal_patrimonio+virg_pat;
	
		if (patrimonio != ""){
			subtotal_patrimonio = subtotal_patrimonio + patrimonio;
			}
		if (pat_apaga != ''){
			subtotal_patrimonio = subtotal_patrimonio.replace(pat_apaga,'');
			}
		
	document.Form.subtotal_patrimonio.value=subtotal_patrimonio;
	document.Form.cd_patrimonio_busca.value="";
	

	/*var oHTTPRequest_pat = createXMLHTTP(); 
    oHTTPRequest_pat.open("post", "../ajax/patrimonio/ajax_patrimonio_lista.asp", true);
    oHTTPRequest_pat.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    oHTTPRequest_pat.onreadystatechange=function(){
     if (oHTTPRequest_pat.readyState==4){
        document.all.divPat_lista.innerHTML = oHTTPRequest_pat.responseText;}}
       oHTTPRequest_pat.send("subtotal_patrimonio=" + Form.subtotal_patrimonio.value);
	   */
	//chama Ajax	
	loadXMLDoc("../ajax/procedimento/ajax_patrimonio_lista.asp?subtotal_patrimonio="+subtotal_patrimonio,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
						document.getElementById("divproc_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		});
 
}
//*********************************************************************************

function soma_otica(u,v,otica_total){
	if (otica_total != ''){
		virg =  ',';
		}
	else{
		virg= ''
	}		
	otica_total = otica_total + virg
		if (u != ""){
			otica_total = otica_total + u;
			}
		if (v != ""){
			otica_total = otica_total.replace(v,'');
			}	
		
	document.Form.otica_res.value=u;
	document.Form.otica_total.value=(otica_total.replace(",,", ","));
	document.Form.cd_otica_busca.value="";
	//document.Form.cd_procedimento_1.value="";
	document.Form.cd_otica_1.value="";
		
/*
     var oHTTPRequest_otica = createXMLHTTP(); 
     oHTTPRequest_otica.open("post", "../ajax/procedimento/ajax_otica_lista.asp", true);
     oHTTPRequest_otica.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_otica.onreadystatechange=function(){
      if (oHTTPRequest_otica.readyState==4){
         document.all.divOtica_lista.innerHTML = oHTTPRequest_otica.responseText;}}
       oHTTPRequest_otica.send("otica_total=" + form.otica_total.value);
	   */
 	loadXMLDoc("../ajax/procedimento/ajax_otica_lista.asp?otica_total="+otica_total,function(){
		  		if (xmlhttp.readyState==4 && xmlhttp.status==200){
						document.getElementById("divOtica_lista").innerHTML = xmlhttp.responseText;
		    		}
		  		});
}
//*********************************************************************************
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
//**********************************************************************************************************************

	


// End -->
</script>

<%cd_user = session("cd_codigo")

pasta_loc = "protocolo\"
arquivo_loc = "protocolo_digitacao.asp"

cd_unidade = request("cd_unidade")
no_protocolo = request("no_protocolo")
no_digito = request("no_digito")

cd_codigo = request("cd_codigo")
tipo = request("tipo")
modo = request("modo")
modalidade = request("modalidade")
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
'elseif action_form <> "" Then
else
	if cd_unidade = "" AND no_protocolo = "" Then
		action_form = "protocolo.asp"
	else
		action_form = "protocolo/acoes/protocolo_acao.asp"
	end if
end if

'*******************************
'*** Libera trava para Datas ***
'*******************************
cd_libera_data = "0"
'*******************************


'modo = "ins"
%>
<!--include file="../js/ajax.js"-->
<!--#include file="../js/ajax2.asp"-->
<!--#include file="../js/protocolo.js"-->

<BODY onload="foco_protocolo();">
<!--#include file="../includes/arquivo_loc.asp"-->
<!--body--><br>
<!--Cabeçalho do formulário -->
	<table border="0" class="textopequeno" align="center" style="collapse-border:collpase;">
	<!--table border="1" cellspacing="0" cellpadding="0" class="textopequeno" align="center" width="300px"-->
		<tr>		
			<td class="txt_cinza" colspan="4">
			<b>Protocolos &raquo; <%=tipo%><br><br></td>
			</td>
		</tr>
		<tr>
			<td colspan="4">
					&nbsp;::&nbsp; <a href="protocolo.asp?tipo=digitacao&modalidade=novo">Novo</a>
					&nbsp;::&nbsp; <a href="protocolo.asp?tipo=digitacao&modalidade=alterar&no_protocolo=<%=no_protocolo%>&cd_unidade=<%=cd_unidade%>&no_digito=<%=no_digito%>&action_form=protocolo/acoes/protocolo_acao.asp">Alterar</a> 
					&nbsp;::&nbsp; <a href="protocolo.asp?tipo=ver">Consultar</a>
					&nbsp;::&nbsp; <a href="protocolo.asp?tipo=digitacao&modalidade=cancelar">Cancelados</a>&nbsp;::&nbsp;
			</td>
		</tr>
		<!--tr id="no_print"><td colspan="4"><img src="imagens/px.gif" alt="" width="10" height="5" border="0"></td></tr-->
		<tr bgcolor="#808080">
			<td align="center" colspan="4">&nbsp;<font size="2" color="white"><b>Dados do Protocolo</b></font></td>
		</tr>
		<form method="post" action="<%=action_form%>" name="Form" id="form" onsubmitzzz="return validar_protocolo(document.Form);" accept-charset="iso-8859-1,utf-8">
		<input type="hidden" name="modo" value="<%=modo%>">
		<input type="hidden" name="tipo" value="<%=tipo%>">
		<input type="hidden" name="modalidade" value="<%=modalidade%>">
		<input type="hidden" name="action_form" value="ok">
		<!--input type="hidden" name="acao" value="<%'=acao%>"-->


<%if cd_unidade = "" AND no_protocolo = "" Then
'********************************************************************
'***** 1.0 envia dados do protocolo para avaliação (ver/editar) *****
'********************************************************************
		if modalidade = "novo" Then 'NOVO
			cor_proto = "ccffcc"
		elseif modalidade = "alterar" Then 'altera
			cor_proto = "ffff99"
		elseif modalidade = "cancelar" then 'Cancela
			cor_proto = "ffefce"
			else
			cor_proto = "f5f5f5"
		end if
%>
		<tr bgcolor="#<%=cor_proto%>">
			<td align="left">&nbsp;Protocolo  <font color="red"><b><%=alerta_unid%></b></font></td>
			<td align="center" colspan="3">&nbsp;
				<input type="hidden" name="acao" value="<%=acao%>">
				<input type="text" name="cd_unidade" id="cd_unidade" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" class="inputs" value="<%=cd_unidade%>">.
				<input type="text" name="no_protocolo" size="6" maxlength="6" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 6)" onFocus="PararTAB(this);" class="inputs" value="<%=protocolo%>">
				<input type="text" name="no_digito" size="1" maxlength="1" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 1)" onFocus="PararTAB(this);" onBlur="validar_protocolo()" class="inputs" value="<%=no_digito%>">				
				<img src="../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="foca_unidade();">
				<img src="../../imagens/blackdot.gif" alt="" width="1" height="1" border="0" onLoad="foca_unidade();">
				<br><%=mensagem%>
			</td>
		</tr>
<%elseif cd_unidade <> "" AND no_protocolo <> "" Then
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
							response.redirect("protocolo.asp?tipo=ver&mensagem=(Unidade Inválida)&protocolo="&no_protocolo&"")
						end if
						
						
'*********************************************
'***** 1.2 Verifica o digito verificador *****
'*********************************************
				
					'Verifica se o nº do protocolo foi gerado pelo Siga.
					if cd_unidade = 2 and no_protocolo < 17222 OR cd_unidade = 3 and no_protocolo < 114145 OR cd_unidade = 11 and no_protocolo < 1476 OR cd_unidade = 14 and no_protocolo < 2955  OR cd_unidade = 15 and no_protocolo < 3102 OR cd_unidade = 19 and no_protocolo < 2412 OR cd_unidade = 20 and no_protocolo < 4138 OR cd_unidade = 22 and no_protocolo < 65 then
						response.write("número gerado pelo SIGA: <b>"&cd_unidade&"."&no_protocolo&"</b>-X")
					Else
					
						verifica_digito = digitov(cd_unidade&"."&proto(no_protocolo))
						
						if right(verifica_digito,1) = no_digito then
							'response.write("número correto!: "&verifica_digito)
							'response.write("número correto!")
						elseif no_digito = "o" then
							'response.write("Pula a verificação do digito")
						else
						response.write("Numero inválido <br>")
						response.write("número correto: "&verifica_digito)
						response.redirect("protocolo.asp?tipo="&tipo&"&mensagem=(Número Inválido)")
						end if
					end if
				
'*******************************************************
'***** 1.3 Verifica se o nº do protocolo já existe *****
'*******************************************************

			xsql ="SELECT * FROM TBL_protocolo WHERE A002_numpro = '"&no_protocolo&"' AND A053_codung = '"&cd_unidade&"'"
					Set rs = dbconn.execute(xsql)
					do while Not rs.EOF
					
						'cd_codigo = rs("a002_numseq")
						codigo = rs("cd_cod_protocolo")
						cd_cod_protocolo = rs("cd_cod_protocolo")
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
						cd_co = rs("cd_co")
							if int(cd_co) = 1 then
								co_ck = "checked"
							end if
					cd_cancelado = rs("cd_cancelado")
					
						if cd_cod_protocolo <> "" Then
							proc_altera = " "'request
						end if
					
							
					rs.movenext
					loop
						
						
						rs.close
						set rs = nothing
						'if cd_codigo <> "" Then%>
							
							<%if tipo = "digitacao" AND cd_cod_protocolo <> "0" and modalidade = "novo" then%>							
							<%some = 0%>
							<!--#include file="protocolo_dig.asp"-->
							
							<%elseif tipo = "digitacao" AND cd_cod_protocolo <> "0" and modalidade = "alterar" then%>							
							<%some = 0%>
							<!--#include file="protocolo_dig.asp"-->
							
							<%elseif tipo = "digitacao" AND cd_cod_protocolo <> "0"  and modalidade = "cancelar" then%>
							<%some = 1%>
							<!--#include file="protocolo_canc.asp"-->
							
							<%elseif tipo = "ver" AND cd_cod_protocolo <> ""  and modalidade =  "" then%>
							<%some = 0%>
							<!--#include file="protocolo_ver.asp"-->
							
							<%elseif tipo = "dvd" AND cd_cod_protocolo <> ""  and modalidade =  "" then%>
							<%some = 0%>
							<!--#include file="protocolo_dvd.asp"-->
							
							<%elseif tipo = "ver" AND cd_cod_protocolo = ""  and modalidade =  "" then%>
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
	<tr bgcolor="#f7f7f7">
		<td colspan="3" align="center" valign="center">
		<%if some <> "1" Then%>
		
		<input type="hidden" name="vazio" value="" id="verproce">
		
		    <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js" ></script>

<script type="text/javascript">


$(document).ready(function(){
$("#cd_procedimento_busca").click(function(){


$("#verproce").val("validado");

});
});








$(document).ready(function(){
    $("#cd_procedimento_busca").blur(function(){
     if($("#verproce").val() == "")
         {
             $("#cd_procedimento_busca").css({"border" : "1px solid #F00", "padding": "2px"});
         }
    });
    $("#enviaprot").click(function(){
     var cont = 0;
     $("#cd_procedimento_busca").each(function(){
         if($("#verproce").val() == "")
             {
                 $("#cd_procedimento_busca").css({"border" : "1px solid #F00", "padding": "2px"});
                 cont++;
             }
        });
     if(cont == 0)
         {
             $("#form").submit();
         }
    });
});
</script>
		
			
			<input type="Button" name="Submit" id="enviaprot" value="ok" onkeypress="detecta_enter(event,'submit','<%=cd_user%>')"> &nbsp; <%'=action_form%>
		<%end if%>
		<a href="protocolo.asp?tipo=<%=tipo%>&modalidade=<%=modalidade%>">Limpar</a></td>
		<td>
			<%if cd_user = 46 then%><img src="../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onclick="javascript:JsProtocoloDelete(<%=codigo%>);"><%end if%>
		</td>
	</tr>
	<tr>
		<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
		<td><img src="../imagens/px.gif" alt="" width="10" height="1" border="0"></td>
	</tr>	
			</form>
	<!--tr><td><img src="../imagens/blackdot.gif" alt="" width="1" height="1" border="0"></td></tr-->
</table><br>

		
<%'rs.close
'set rs = nothing
















%>

<br>
</BODY>
</HTML>
<SCRIPT LANGUAGE="javascript">
			
function JsProtocoloDelete(cod){
		if (confirm ("Tem certeza que deseja excluir esse protocolo?")){
			document.location=('protocolo/acoes/protocolo_acao.asp?cd_codigo='+cod+'&acao=apaga_protocolo_inteiro');
		}
	} 
</SCRIPT>
			
