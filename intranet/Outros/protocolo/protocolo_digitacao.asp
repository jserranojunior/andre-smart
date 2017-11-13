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

if cd_codigo = "" AND tipo = "ins" Then
acao = "inserir"
Elseif cd_codigo <> "anything" AND tipo = "ver" then
acao = "editar"
end if
%>


<BODY onLoad="foco()"><br>
<%
'*******************************************************
'***** 1.1 verifica se já existe no Banco de Dados *****
'*******************************************************

		'xsql ="SELECT * FROM TBL_PROTOCOLO WHERE cd_protocolo = '"&cd_protocolo&"'"
		xsql ="SELECT * FROM TBL_PROTOCOLO WHERE A002_numpro = '"&cd_protocolo&"' AND A053_codung='"&cd_unidade&"'"
		Set rs = dbconn.execute(xsql)
			Do while NOT rs.EOF
			
			'cd_prot = rs("cd_protocolo")
			'cd_unidade = rs("cd_unidade")
			'nm_nome = rs("nm_nome")
			'cd_idade = rs("cd_idade")
			'cd_registro = rs("cd_paciente")
			'cd_convenio = rs("cd_convenio")
			'nm_sexo = rs("nm_sexo")
			'nm_cirug_realizada = rs("nm_cirug_realizada")
			'dt_procedimento = rs("dt_procedimento")
			'nm_agenda = rs("nm_agenda")
			'dt_hr_agenda = rs("dt_hr_agenda")
			'dt_inicio = rs("dt_inicio")
			'dt_fim = rs("dt_fim")
			'cd_crm = rs("cd_medico")
			'cd_rack = rs("cd_rack")
			'cd_especialidade = rs("cd_especialidade")
			'cd_procedimento = rs("cd_procedimento")
			
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
			cd_med = rs("A055_codmed")
			cd_rack = rs("a056_codrac")
			cd_especialidade = rs("A057_codesp")
			'cd_procedimento = rs("cd_procedimento")
			
'			nm_sobrenome = rs("nm_sobrenome")
			rs.movenext
			Loop				

'**********************************************
'***** 1.2 verifica se a unidade é válida *****
'**********************************************

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

'**********************************************
'***** 1.3 define o destino do formulário *****
'**********************************************

If     tipo = "ins" AND cd_protocolo = "" AND cd_unid = "" Then
			action_form = "protocolo.asp"
Elseif tipo = "ins" AND cd_protocolo <> "" AND cd_unid <> "" Then
			action_form = "protocolo/protocolo_acao.asp"
Elseif tipo = "ver" AND cd_protocolo = "" AND cd_unid = "" Then
			action_form = "protocolo.asp"
Elseif tipo = "ver" AND cd_protocolo <> "" AND  cd_unid <> "" Then
			action_form = "protocolo/protocolo_acao.asp"
End if
%>

<table align="center" border="1" cellspacing="0" cellpadding="0"><tr><td>
<!--Cabeçalho do formulário -->
	<table width="650" border="1" cellspacing="1" cellpadding="" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
		<tr>		
			<td class="txt_cinza" colspan="100%">
			<b>Protocolos &raquo; - <font color="red">(<%=acao%>:<%=action_form%>)</font></b><br><br></td>
			</td>
		</tr>
		<tr>			
				
		<tr bgcolor="#808080">
			<td  colspan="100%" align="center" class="textopadrao">&nbsp;<font color="white"><b>Dados do Protocolo</b></font></td>
		</tr>
		<tr bgcolor="#f5f5f5">
			<td align="left">&nbsp;Protocolo  <font color="red"><b><%=alerta_unid%></b></font></td>
			<form method="post" action="<%=action_form%>?tipo=<%=tipo%>" name="Form" id="form">	
			<%
			'*******************************************************
			'***** 2.0 envia dados do protocolo para avaliação *****
			'*******************************************************
			if cd_protocolo = "" OR cd_unid = "" Then
				mostra = "0"%>
					<td>&nbsp;
									<input type="text" name="cd_unidade" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" class="inputs" value="">.
									<input type="text" name="cd_protocolo" size="6" maxlength="6" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 6)" onFocus="PararTAB(this);" class="inputs" value="<%=cd_protocolo%>">
									<input type="text" name="cd_digito" size="1" maxlength="1" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 1)" onFocus="PararTAB(this);" onBlur="document.form.submit()" class="inputs" value="<%=cd_digito%>">&nbsp;
									<input type="hidden" name="tipo" value="ins">
					</td>
			<%'end if%>						
									
			<%
			'***********************************************
			'***** 2.1 avisa que o protocolo já existe *****
			'***********************************************
			elseif cd_prot = int(cd_protocolo) AND tipo = "ins" Then
				mostra = "0"%>
					<td colspan="2">&nbsp;O protocolo digitado já existe!</td>
			<%
			'**************************************************
			'***** 2.2 Mostra Nº do protocolo e a Unidade *****
			'**************************************************
			elseif cd_prot = "" AND tipo = "ins" OR cd_prot <> "" AND tipo = "ver" Then
				mostra="1"%>		
					<td>&nbsp;<%=cd_unidade%>.<b><%=cd_protocolo%></b><%if cd_digito <> "" Then%><%="-"&cd_digito%><%end if%></td>		
			
					<td align="left">&nbsp;Unidade</td>
					<td align="left" colspan="2">&nbsp;<%xsql ="SELECT * FROM TBL_unidades WHERE cd_codigo = '"&cd_unidade&"'"
					Set rs = dbconn.execute(xsql)
					do while Not rs.EOF
					response.write("<b>"&rs("nm_unidade")&"</b>")
					rs.movenext
					loop%>		
					</td>
			<%
			'*********************************************************
			'***** 2.3 avisa que não existe o protocolo digitado *****
			'*********************************************************
			Elseif cd_prot = "" AND tipo = "ver" Then
				mostra="0"%>
					<td colspan="2">&nbsp;O protocolo digitado não existe!</td>
			<%end if%>	
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha-->*:<%=cd_prot%>:*</td></tr>
			<%
			'****************************************
			'***** Dados do protocolo - 2ª fase *****
			'****************************************
			'if tipo="ver" AND cd_prot<>"" OR tipo="ins" AND cd_prot<>"" Then
			if mostra = "1" Then%>
			
			
		<tr bgcolor="#808080">
			<td  colspan="100%" align="center" class="textopadrao"><font color="white"><b>&nbsp;Identificação do Paciente</b></font></td>
		</tr>
		<tr>
			<td align="left" colspan="3">&nbsp;Nome</td>
			<td align="left"  colspan="2">&nbsp;Idade</td>
		</tr>
		<tr>
			<td align="left" colspan="3">&nbsp;<input type="text" name="nm_nome" size="65" maxlength="60" class="inputs" value="<%=nm_nome%>"></td>
			<td align="left">&nbsp;<input type="text" name="cd_idade" size="5" maxlength="3" class="inputs" value="<%=cd_idade%>"></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha-->
					<input type="hidden" name="acao" value="<%=acao%>">
					<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
					<input type="hidden" name="cd_protocolo" value="<%=cd_protocolo%>">
					<!--input type="hidden" name="cd_digito" value="<%=cd_digito%>"-->
					</td></tr>
		<tr>
			<td align="left">&nbsp;Registro</td>
			<td align="left" colspan="2">&nbsp;Convênio: </td>
			<td align="left">&nbsp;Sexo: </td>
		</tr>
		<tr>
			<td align="left">&nbsp;<input type="text" name="nm_registro" size="10" maxlength="15" class="inputs" value="<%=cd_registro%>">&nbsp;&nbsp;
			<td align="left" valign="top" colspan="2"><!--input type="text" name="cd_convenio" size="3" class="inputs">
								<select name="cd_convenio_1" class="inputs">
								<option value=""></option>
								<option value="1">Sul America Empresas</option>
								</select-->
								<%strsql ="Select * from TBL_convenio order by nm_convenio"
						  		Set rs_conv = dbconn.execute(strsql)%>
							<input type="text" name="cd_convenio" size="3" maxlength="4" class="inputs" value="<%=cd_convenio%>">	
							<select name="cd_convenio_1" class="inputs"><option value="">
								<%Do while Not rs_conv.EOF%><%cd_conv = rs_conv("cd_codigo")%>
								<%if cd_convenio=cd_conv Then%><%ck_conv="selected"%><%else ck_conv=""%><%end if%>
								<option value="<%=rs_conv("cd_codigo")%>" <%=ck_conv%>><%=rs_conv("nm_convenio")%></option><%rs_conv.movenext
						 		Loop%></td>
			<td align="left"><select name="nm_sexo" size="1" class="inputs">
							 <option value=""></option>
							 <%if nm_sexo="M" Then%><%ck_m = "selected"%>
							 <%Elseif nm_sexo="F" Then%><%ck_f = "selected"%>
							 <%end if%>
							 <option value="M" <%=ck_m%>>M</option>
							 <option value="F" <%=ck_f%>>F</option>&nbsp;&nbsp;</select></td>
			
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
					<%if nm_cirug_realizada="S" Then%><%ck_real_s = "selected"%>
					<%Elseif nm_cirug_realizada="N" Then%><%ck_real_n = "selected"%>
					<%end if%>
					<option value="S" <%=ck_real_s%>>SIM</option>
					<option value="N" <%=ck_real_n%>>NÃO</option>
				</select></td>
				<% 			if dt_procedimento <> "" Then
								dt_proced_dia = DAY(dt_procedimento)
								dt_proced_mes = MONTH(dt_procedimento)
								dt_proced_ano = YEAR(dt_procedimento)
							end if%>
			<td colspan="2"><input type="text" name="dt_dia_cirurgia" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_proced_dia%>">
							<input type="text" name="dt_mes_cirurgia" size="2" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_proced_mes%>">
							<input type="text" name="dt_ano_cirurgia" size="4" maxlength="4" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_proced_ano%>"></td>
			<td><select name="nm_agenda_cirurgia" class="inputs">
					<option value=""></option>
					<%if nm_agenda="A" Then%><%ck_agd_a = "selected"%>
					<%Elseif nm_agenda="E" Then%><%ck_agd_e = "selected"%>
					<%Elseif nm_agenda="N" Then%><%ck_agd_n = "selected"%>
					<%Elseif nm_agenda="P" Then%><%ck_agd_p = "selected"%>
					<%Elseif nm_agenda="U" Then%><%ck_agd_u = "selected"%>
					<%end if%>
					<option value="A" <%=ck_agd_a%>>A seguir</option>
					<option value="E" <%=ck_agd_e%>>Empréstimo</option>
					<option value="N" <%=ck_agd_n%>>Normal</option>
					<option value="P" <%=ck_agd_p%>>Pós Agendada</option>
					<option value="U" <%=ck_agd_u%>>Urgência</option>
					
				</select></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr>
			<td>Hora Agendada</td>
			<td colspan="2">Início</td>
			<td>Término</td>
		</tr>
		<%
		if dt_hr_agenda <> "" Then
			dt_min_agenda = minute(dt_hr_agenda)
			dt_hr_agenda = hour(dt_hr_agenda)
		end if
		if dt_inicio <> "" Then
			dt_min_inicio = minute(dt_inicio)
			dt_hr_inicio = hour(dt_inicio)
		end if
		if dt_fim <> "" Then
			dt_min_fim = minute(dt_fim)
			dt_hr_fim = hour(dt_fim)
		end if
		%>
		<tr>
			<td><input type="text" size="2" name="dt_hora_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_hr_agenda%>"> :
				<input type="text" size="2" name="dt_minuto_agenda" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_min_agenda%>"></td>
			<td colspan="2"><input type="text" size="2" name="dt_hora_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_hr_inicio%>"> :
				<input type="text" size="2" name="dt_minuto_inicio" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_min_inicio%>"></td>
			<td><input type="text" size="2" name="dt_hora_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_hr_fim%>"> :
				<input type="text" size="2" name="dt_minuto_fim" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" maxlength="3" class="inputs" value="<%=dt_min_fim%>"></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr bgcolor="#808080">
			<td  colspan="100%" align="center" class="textopadrao"><font color="white"><b>&nbsp;Procedimento Cirúrgico</b></font></td>
		</tr>
		<tr>
			<td colspan="6">CRM</td>
		</tr>
		<tr>				<%'sql_med ="Select * from TBL_medicos Where cd_codigo = '"&A055_codmed&"' order by nm_medico"
						  		'Set rsmed = dbconn.execute(sql_med)
								'if not rsmed.EOF Then
								'cd_crm = rsmed("cd_crm")
								'end if%>
			<td colspan="6"><input type="text" name="cd_crm" class="inputs" size="10" maxlength="10" value="<%=cd_crm%>">
							<%strsql ="Select * from TBL_medicos order by nm_medico"
						  		Set rs_med = dbconn.execute(strsql)
								if not rs_med.EOF Then
								cd_crm = rs_med("cd_crm")
								end if%>
							<select name="cd_crm_1" class="inputs">
								<option value=""></option>
								<%Do while Not rs_med.EOF%><%cd_medico = rs_med("cd_codigo")%>
								<%if cd_medico =cd_med Then%><%ck_med="selected"%><%else ck_med=""%><%end if%>
								<option value="<%=rs_med("cd_codigo")%>" <%=ck_med%>><%=rs_med("nm_medico")%></option><%rs_med.movenext
						 		Loop%></td>
							</select></td>
		</tr>
		<tr>
			<td>Rack</td>
			<td>Especialidade</td>
		</tr>
		<tr>
			<td><input type="text" name="cd_rack" size="3" class="inputs" value="<%=cd_rack%>">
							<select name="cd_rack_1" class="inputs">
								<option value=""></option>
								<%strsql ="Select * from TBL_Rack"' order by nm_rack"
						  		Set rs_rack = dbconn.execute(strsql)%>
								<%Do while Not rs_rack.EOF%><%cd_rk = rs_rack("cd_codigo")%>
								<%if cd_rack=int(cd_rk) Then%><%ck_rack="selected"%><%else ck_rack=""%><%end if%>
								<option value="<%=rs_rack("cd_codigo")%>" <%=ck_rack%>><%=rs_rack("nm_rack")%></option><%rs_rack.movenext
						 		Loop%></td>
							</select></td>
			<td><input type="text" name="cd_especialidade" size="3" class="inputs" value="<%=cd_especialidade%>">
							<select name="cd_especialidade_1" class="inputs">
								<option value=""></option>
								<%strsql ="Select * from TBL_espec order by nm_especialidade"
						  		Set rs_espec = dbconn.execute(strsql)%>
								<%Do while Not rs_espec.EOF%><%cd_espec = rs_espec("cd_codigo")%>
								<%if cd_especialidade=cd_espec Then%><%ck_espec="selected"%><%else ck_espec=""%><%end if%>
								<option value="<%=rs_espec("cd_codigo")%>" <%=ck_espec%>><%=rs_espec("nm_especialidade")%></option><%rs_espec.movenext
						 		Loop%>
							</select></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<tr>
			<td colspan="100%">Procedimento</td>
		</tr>
		<tr>
			<td colspan="100%"><input type="text" name="cd_procedimento" size="10" maxlength="10" class="inputs">
							<select name="cd_procedimento_1" class="inputs">
								<option value=""></option>
								<%strsql ="Select * from TBL_procedimento order by nm_procedimento"
						  		Set rs_proc = dbconn.execute(strsql)%>
								<%Do while Not rs_proc.EOF%><%cd_proc = rs_proc("cd_codigo")%>
								<option value="<%=rs_proc("cd_codigo")%>" <%=ck_proc%>><%=left(rs_proc("nm_procedimento"),25)%> (<%=rs_proc("cd_codigo")%>)</option><%rs_proc.movenext
						 		Loop%>
							</select>
							<input type="button" name="somar" value="Inserir" onclick="soma(document.form.cd_procedimento.value,document.form.cd_procedimento_1.value,document.form.procedimentos.value)">
			</td>
		</tr>
		</tr>
			<td colspan="2">&nbsp;<textarea cols="30" rows="5" name="procedimentos"><%=cd_procedimento%></textarea>
							<input type="hidden" name="res" size="5">
							</td>
							<%strsql ="Select * from view_protocolo_procedimento where A002_numseq = '"&cd_prot&"'"'order by nm_procedimento"
						  	Set rs_proc2 = dbconn.execute(strsql)%>
								
			<td colspan="2">
				<%do while not rs_proc2.EOF%>
				<%=response.write("<b>"&rs_proc2("A058_codpro")&"</b> - "& rs_proc2("nm_procedimento"))%><br>
				<%rs_proc2.movenext
				loop
				%></td>
		</tr>
		<tr><td colspan="100%">&nbsp;<!--Pula linha--></td></tr>
		<%else%>
	<%end if%>

	<tr>
		<td colspan="100%" align="center" valign="center">
		<%'if cd_prot = "0" OR cd_prot = "" Then%>
			<input type="submit" name="Submit" value="ok"> &nbsp; 
		<%'end if%>
		<a href="protocolo.asp?tipo=<%=tipo%>">Limpar</a></td>
	</tr>
			</form>
</table>
</td></tr></table><br>
<br>
</BODY>
</HTML>


