<!--#include file="../includes/inc_open_connection.asp"--> 

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
			virg =  ',\r\n';
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
	document.form.procedimentos.value=procedimentos;
	document.form.cd_procedimento.value="";
	document.form.cd_procedimento_1.value="";
}
// -->	
</script>

<%cd_unidade = request("cd_unidade")
cd_protocolo = request("cd_protocolo")
cd_digito = request("cd_digito")
cd_codigo = request("cd_codigo")
tipo = request("tipo")

if cd_codigo = "" Then
acao = "inserir"
Else
acao = "editar"
end if
%>


<BODY onLoad="foco()" ><br>
<%
if tipo = "ins" Then
'***** verifica se já existe no Banco de Dados *****
		xsql ="SELECT cd_protocolo FROM TBL_PROTOCOLO WHERE cd_protocolo = '"&cd_protocolo&"'"
		Set rs = dbconn.execute(xsql)
			Do while NOT rs.EOF
			cd_prot = rs("cd_protocolo")
			rs.movenext
			Loop				
'***** verifica se a unidade é válida *****
		if cd_unidade <> "" Then
			xsql ="SELECT * FROM TBL_UNIDADES WHERE cd_codigo = '"&cd_unidade&"'"
			Set rs = dbconn.execute(xsql)
				Do while NOT rs.EOF
				cd_unid = rs("cd_codigo")
				rs.movenext
				Loop
					if cd_unid = "" Then
						alerta_unid = "Unidade Inválida"
					end if
		end if
'***** ****** ******

If cd_protocolo = "" AND cd_unid = "" Then
	action_form = "protocolo.asp"
Elseif cd_protocolo <> "" AND cd_unid <> "" Then
	action_form = "protocolo/protocolo_acao.asp"
End if
%>
 
<table align="center" border="1" cellspacing="0" cellpadding="0"><tr><td>
<table width="550" border="1" cellspacing="1" cellpadding="" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
	<tr>		
		<td class="txt_cinza" colspan="100%">
		<b>Protocolos &raquo; - <font color="red">Digitação<%=acao%></font></b><br><br></td>
		</td>
	</tr>
	<tr>			
			
	<tr bgcolor="#808080">
		<td  colspan="100%" align="center" class="textopadrao">&nbsp;<font color="white"><b>Dados do Protocolo</b></font></td>
	</tr>
	<tr bgcolor="#f5f5f5">
		<td align="left">&nbsp;Protocolo  <font color="red"><b><%=alerta_unid%></b></font></td>
		<form method="post" action="<%=action_form%>" name="Form" id="form" onsubmit="return false">	
<%if cd_protocolo = "" OR cd_unid = "" Then%>
		<td>&nbsp;
						<input type="text" name="cd_unidade" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" class="inputs" value="">.
						<input type="text" name="cd_protocolo" size="6" maxlength="6" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 6)" onFocus="PararTAB(this);" class="inputs" value="<%=cd_protocolo%>">-
						<input type="text" name="cd_digito" size="1" maxlength="1" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 1)" onFocus="PararTAB(this);" onBlur="document.form.submit()" class="inputs" value="<%=cd_digito%>">&nbsp;
						<input type="hidden" name="tipo" value="ins">
						</td>
<%elseif cd_prot <> "" Then%>
		<td colspan="2">&nbsp;O protocolo digitado já existe!</td>
<%elseif cd_prot = "" Then%>		
		<td>&nbsp;<%=cd_unidade%>.<b><%=cd_protocolo%></b>-<%=cd_digito%></td>		

		<td align="left">&nbsp;Unidade</td>
		<td align="left" colspan="2">&nbsp;<%xsql ="SELECT * FROM TBL_unidades WHERE cd_codigo = '"&cd_unidade&"'"
		Set rs = dbconn.execute(xsql)
		do while Not rs.EOF
		response.write("<b>"&rs("nm_unidade")&"</b>")
		rs.movenext
		loop%>		
		</td>

	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>

	
<%if cd_protocolo <> "" OR cd_unidade <> "" Then%>
	<tr bgcolor="#808080">
		<td  colspan="100%" align="center" class="textopadrao"><font color="white"><b>&nbsp;Identificação do Paciente</b></font></td>
	</tr>
	<tr>
		<td align="left" colspan="3">&nbsp;Nome</td>
		<td align="left"  colspan="2">&nbsp;Idade</td>
	</tr>
	<tr>
		<td align="left" colspan="3">&nbsp;<input type="text" name="nm_nome" size="65" maxlength="60" class="inputs"></td>
		<td align="left">&nbsp;<input type="text" name="cd_idade" size="5" maxlength="3" class="inputs"></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha-->
				<input type="hidden" name="acao" value="<%=acao%>">
				<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
				<input type="hidden" name="cd_protocolo" value="<%=cd_protocolo%>">
				<input type="hidden" name="cd_digito" value="<%=cd_digito%>">
				</td></tr>
	<tr>
		<td align="left">&nbsp;Registro</td>
		<td align="left" colspan="2">&nbsp;Convênio: </td>
		<td align="left">&nbsp;Sexo: </td>
	</tr>
	<tr>
		<td align="left">&nbsp;<input type="text" name="nm_registro" size="10" maxlength="15" class="inputs">&nbsp;&nbsp;
		<td align="left" valign="top" colspan="2"><!--input type="text" name="cd_convenio" size="3" class="inputs">
							<select name="cd_convenio_1" class="inputs">
							<option value=""></option>
							<option value="1">Sul America Empresas</option>
							</select-->
							<%strsql ="Select * from TBL_convenio order by nm_convenio"
					  		Set rs_conv = dbconn.execute(strsql)%>
						<input type="text" name="cd_convenio" size="3" maxlength="4" class="inputs">	
						<select name="cd_convenio_1" class="inputs"><option value="">
							<%Do while Not rs_conv.EOF%><%cd_convenio = rs_conv("cd_codigo")%>
							<%'if nm_fornecedor=fornecedor Then%><%'ck_forn="selected"%><%'else ck_forn=""%><%'end if%>
							<option value="<%=rs_conv("cd_codigo")%>" <%=ck_conv%>><%=rs_conv("nm_convenio")%></option><%rs_conv.movenext
					 		Loop%></td>
		<td align="left"><select name="nm_sexo" size="1" class="inputs">
						 <option value=""></option>
						 <option value="M">M</option>
						 <option value="F">F</option>&nbsp;&nbsp;</select></td>
		
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr bgcolor="#808080">
		<td  colspan="100%" align="center" class="textopadrao"><font color="white"><b>&nbsp;Dados Gerais</b></font></td>
	</tr>
	<tr>
		<td>Cirurgia Realizada?</td>
		<td colspan="2">Data</td>
		<td>Agenda</td>
	</tr>
	<tr>
		<td><select name="nm_cirug_realizada" class="inputs">
				<option value=""></option>
				<option value="S">SIM</option>
				<option value="N">NÃO</option>
			</select></td>
		<td colspan="2"><input type="text" name="dt_dia_cirurgia" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs">
						<input type="text" name="dt_mes_cirurgia" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs">
						<input type="text" name="dt_ano_cirurgia" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" maxlength="3" class="inputs"></td>
		<td><select name="nm_agenda_cirurgia" class="inputs">
				<option value=""></option>
				<option value="A">A seguir</option>
				<option value="E">Empréstimo</option>
				<option value="N">Normal</option>
				<option value="P">Pós Agendada</option>
				<option value="U">Urgência</option>
				
			</select></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr>
		<td>Hora Agendada</td>
		<td colspan="2">Início</td>
		<td>Término</td>
	</tr>
	<tr>
		<td><input type="text" size="2" name="dt_hora_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs"> :
			<input type="text" size="2" name="dt_minuto_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs"></td>
		<td colspan="2"><input type="text" size="2" name="dt_hora_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs"> :
			<input type="text" size="2" name="dt_minuto_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs"></td>
		<td><input type="text" size="2" name="dt_hora_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs"> :
			<input type="text" size="2" name="dt_minuto_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs"></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr bgcolor="#808080">
		<td  colspan="100%" align="center" class="textopadrao"><font color="white"><b>&nbsp;Procedimento Cirúrgico</b></font></td>
	</tr>
	<tr>
		<td colspan="6">CRM</td>
	</tr>
	<tr>
		<td colspan="6"><input type="text" name="cd_crm" class="inputs" size="10" maxlength="10">
						<select name="cd_crm_1" class="inputs">
							<option value=""></option>
							<%strsql ="Select * from TBL_medicos order by nm_medico"
					  		Set rs_med = dbconn.execute(strsql)%>
							<%Do while Not rs_med.EOF%><%cd_medico = rs_med("cd_codigo")%>
							<%'if nm_fornecedor=fornecedor Then%><%'ck_forn="selected"%><%'else ck_forn=""%><%'end if%>
							<option value="<%=rs_med("cd_codigo")%>" <%=ck_med%>><%=rs_med("nm_medico")%></option><%rs_med.movenext
					 		Loop%></td>
						</select></td>
	</tr>
	<tr>
		<td>Rack</td>
		<td>Especialidade</td>
	</tr>
	<tr>
		<td><input type="text" name="cd_rack" size="3" class="inputs">
						<select name="cd_rack_1" class="inputs">
							<option value=""></option>
							<%strsql ="Select * from TBL_Rack order by nm_rack"
					  		Set rs_rack = dbconn.execute(strsql)%>
							<%Do while Not rs_rack.EOF%><%cd_rack = rs_rack("cd_codigo")%>
							<%'if nm_fornecedor=fornecedor Then%><%'ck_forn="selected"%><%'else ck_forn=""%><%'end if%>
							<option value="<%=rs_rack("cd_codigo")%>" <%=ck_rack%>><%=rs_rack("nm_rack")%></option><%rs_rack.movenext
					 		Loop%></td>
						</select></td>
		<td><input type="text" name="cd_especialidade" size="3" class="inputs">
						<select name="cd_especialidade_1" class="inputs">
							<option value=""></option>
							<%strsql ="Select * from TBL_espec order by nm_especialidade"
					  		Set rs_espec = dbconn.execute(strsql)%>
							<%Do while Not rs_espec.EOF%><%cd_espec = rs_espec("cd_codigo")%>
							<%'if nm_fornecedor=fornecedor Then%><%'ck_forn="selected"%><%'else ck_forn=""%><%'end if%>
							<option value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option><%rs_espec.movenext
					 		Loop%>
						</select></td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
	<tr>
		<td>Procedimento</td>
	</tr>
	<tr>
		<td colspan="2"><input type="text" name="cd_procedimento" size="10" maxlength="10" class="inputs">
						<select name="cd_procedimento_1" class="inputs">
							<option value=""></option>
							<%strsql ="Select * from TBL_procedimento order by nm_procedimento"
					  		Set rs_proc = dbconn.execute(strsql)%>
							<%Do while Not rs_proc.EOF%><%cd_procedimento = rs_proc("cd_codigo")%>
							<option value="<%=rs_proc("cd_codigo")%>" <%=ck_proc%>><%=rs_proc("cd_codigo")%> - <%=left(rs_proc("nm_procedimento"),25)%></option><%rs_proc.movenext
					 		Loop%>
						</select>
						
		</td>
	</tr>
	</tr>
		<td colspan="4">&nbsp;<textarea cols="50" rows="10" name="procedimentos" readonly></textarea>
						<input type="hidden" name="res" size="5">
						<input type="button" name="somar" value="Inserir" onclick="soma(document.form.cd_procedimento.value,document.form.cd_procedimento_1.value,document.form.procedimentos.value)">
						
						
						</td>
	</tr>
	<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
<%end if%>
<%end if%>
	<tr>
		<td colspan="100%" align="center" valign="center">
		<%if cd_prot = "0" OR cd_prot = "" Then%>
			<input type="submit" name="Submit" value="ok"> &nbsp; 
		<%end if%>
		<a href="protocolo.asp?tipo=ins">Limpar</a></td>
	</tr>
				
</table>
</td></tr></table><br>
<br>
	<%Elseif tipo = "ver" then%>

<table cellspacing="0" cellpadding="0" border="1" align="center">
	<tr>
		<td>
			<table width="550" border="1" cellspacing="1" cellpadding="" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
				<tr>		
					<td class="txt_cinza" colspan="100%">
					<b>Protocolos &raquo; - <font color="red">Ver, Atualizar</font></b><br><br></td>
					</td>
				</tr>
				<tr>			
				<tr bgcolor="#808080">
					<td  colspan="100%" align="center" class="textopadrao">&nbsp;<font color="white"><b>Informe o nº do Protocolo</b></font></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="left">&nbsp;Protocolo</td>
						<form name="ver" id="ver" method="post" action="protocolo.asp">	
					<td>&nbsp;
						<input type="text" name="cd_unid_1" size="2" maxlength="2" class="inputs" value="">.
						<input type="text" name="cd_prot_1" size="6" maxlength="6" class="inputs" value="<%=cd_protocolo%>">-
						<input type="text" name="cd_dig_1" size="1" maxlength="1" class="inputs" value="<%=cd_digito%>">&nbsp;
						<input type="hidden" name="tipo" value="ver">
						<input type="submit" name="submit" value="ok">
						</form>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
	
<%end if%>	
</BODY>
</HTML>


