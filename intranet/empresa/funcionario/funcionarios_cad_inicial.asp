<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%
session_cd_usuario = session("cd_codigo")
cd_user = session_cd_usuario
usuario_restrito = int("65")

strcod = request("cod")
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")


If strcod <> "" And strmsg = "" Then
	
	xsql ="up_Funcionario_Lista_por_cd_codigo @cd_codigo="&strcod&""
	Set rs = dbconn.execute(xsql)
		If NOT rs.EOF Then
		str_funcionario = rs("cd_codigo")
		strmatricula = rs("cd_matricula") 
		strnome = rs("nm_nome")
		strfoto = rs("nm_foto")
		strdt_nascimento = rs("dt_nascimento")
			strdia_nascimento = zero(day(strdt_nascimento))
			strmes_nascimento = zero(month(strdt_nascimento))
			strano_nascimento = year(strdt_nascimento)
		strnm_rg = rs("nm_rg")
		strdt_rg_emissao = rs("dt_rg_emissao")
		strnm_rg_expedidor = rs("nm_rg_expedidor")
		strnm_cpf = rs("nm_cpf")
		strnm_tit_eleitor = rs("nm_tit_eleitor")
		strnm_reservista = rs("nm_reservista")
		strmae = rs("nm_mae")
		strpai = rs("nm_pai")
		strnm_endereco = rs("nm_endereco")
		strnr_numero = rs("nr_numero")
		strnm_complemento = rs("nm_complemento")
		strnm_cidade = rs("nm_cidade")
		strnm_estado = rs("nm_estado")
		strnm_cep = rs("nm_cep")
		strbairro = rs("nm_bairro")
		str_nacionalidade = rs("nm_nacionalidade")
		str_pis = rs("cd_pis")
		str_ctps = rs("cd_ctps")
		str_ctps_serie = rs("cd_ctps_serie")
		str_estado_civil = rs("cd_estado_civil")
		
		str_assistencia_medica_matricula = rs("cd_assistencia_medica_matricula")
		str_assistencia_medica_coparticipacao = rs("nr_assistencia_medica_coparticipacao")
		str_assistencia_odonto_matricula = rs("cd_assistencia_odonto_matricula")
		str_assistencia_odonto_coparticipacao = rs("nr_assistencia_odonto_coparticipacao")
		str_vr = rs("cd_vr")
		str_vale_alimentacao = rs("cd_vale_alimentacao")
		str_cmt_bom = rs("cd_cmt_bom")
		str_sptrans = rs("cd_sptrans")
		
		str_banco = rs("nm_banco")
		str_banco_ag = rs("cd_banco_ag")
		str_banco_conta = rs("cd_banco_conta")
		
		cd_numreg = rs("cd_numreg")
		str_estado_civil = rs("cd_estado_civil")
		str_conjuge = rs("nm_conjuge")
		str_conjuge_nasc = rs("dt_conjuge_nasc")
		str_cd_plano_saude_conj = rs("cd_plano_saude_conj")
		str_naturalidade = rs("nm_naturalidade")
		strnr_tit_eleitor_zona = rs("nr_tit_eleitor_zona")
		strnr_tit_eleitor_seccao = rs("nr_tit_eleitor_seccao")
		strnm_reservista_cat = rs("nm_reservista_cat")
		strnr_hepatite_b = rs("nr_hepatite_b")
		strnr_dupla_adulto = rs("nr_dupla_adulto")
		strdt_dupla_adulto_validade = rs("dt_dupla_adulto_validade")
		strnr_scr = rs("nr_scr")
		strdt_exame_medico = rs("dt_exame_medico")
		strnr_caract_altura = rs("nr_caract_altura")
		strnr_caract_peso = rs("nr_caract_peso")
		strnr_caract_manequim = rs("nr_caract_manequim")
		strnr_caract_calcado = rs("nr_caract_calcado")
		strnm_parente = rs("nm_parente")
		strcd_parente_parentesco = rs("cd_parente_parentesco")
		strcd_residencia_tipo = rs("cd_residencia_tipo")
		strcd_residencia_financ = rs("cd_residencia_financ")
		strcd_residencia_reside = rs("cd_residencia_reside")
		strcd_veiculo_tipo = rs("cd_veiculo_tipo")
		strcd_veiculo_financ = rs("cd_veiculo_financ")
		
		str_conspro = rs("cd_conspro")
		str_numreg = rs("cd_numreg")
		str_rgpro_concessao = rs("rgpro_concessao")
		str_rgpro_status = rs("cd_rgpro_status")
		str_rgpro = rs("obs_rgpro")
		str_dt_rgproinscr = rs("dt_rgproinscr")
		str_dt_rgpropag = rs("dt_rgpropag")
		str_dt_rgproval = rs("dt_rgproval")
		
		str_endereco_alt = rs("nm_endereco_alt")
		
		str_sexo = rs("cd_sexo")		
		
		End If
End if
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
<script language="JavaScript">
function validar_cad(shipinfo){
	if (shipinfo.nm_nome.value.length==""){window.alert ("O nome deve ser preenchido.");shipinfo.nm_nome.focus();return (false);}
	if (shipinfo.cd_sexo.value.length==""){window.alert ("O sexo deve ser preenchido.");shipinfo.cd_sexo.focus();return (false);}
	if (shipinfo.nm_endereco.value.length==""){window.alert ("O endereço deve ser preenchido.");shipinfo.nm_endereco.focus();return (false);}
	if (shipinfo.nr_numero.value.length==""){window.alert ("O número deve ser preenchido.");shipinfo.nr_numero.focus();return (false);}
	if (shipinfo.nm_bairro.value.length==""){window.alert ("O bairro deve ser preenchido.");shipinfo.nm_bairro.focus();return (false);}
	if (shipinfo.nm_cep.value.length==""){window.alert ("O CEP deve ser preenchido.");shipinfo.nm_cep.focus();return (false);}
	if (shipinfo.nm_cidade.value.length==""){window.alert ("A cidade deve ser preenchida.");shipinfo.nm_cidade.focus();return (false);}
	if (shipinfo.nm_estado.value.length==""){window.alert ("O estado deve ser preenchido.");shipinfo.nm_estado.focus();return (false);}
	if (shipinfo.dia_nascimento.value.length==""){window.alert ("Informe a data de nascimento.");shipinfo.dia_nascimento.focus();return (false);}
	if (shipinfo.mes_nascimento.value.length==""){window.alert ("Informe a data de nascimento.");shipinfo.mes_nascimento.focus();return (false);}
	if (shipinfo.ano_nascimento.value.length==""){window.alert ("Informe a data de nascimento.");shipinfo.ano_nascimento.focus();return (false);}
	if (shipinfo.nm_nacionalidade.value.length==""){window.alert ("informe a nacionalidade.");shipinfo.nm_nacionalidade.focus();return (false);}
	if (shipinfo.nm_naturalidade.value.length==""){window.alert ("Informe a cidade de nascimento.");shipinfo.nm_naturalidade.focus();return (false);}
	if (shipinfo.cd_estado_civil.value.length==""){window.alert ("Informe o estado civil.");shipinfo.cd_estado_civil.focus();return (false);}
	if (shipinfo.nm_mae.value.length==""){window.alert ("Informe o nome da mãe.");shipinfo.nm_mae.focus();return (false);}
	if (shipinfo.nm_pai.value.length==""){window.alert ("Informe o nome do pai.");shipinfo.nm_pai.focus();return (false);}
	if (shipinfo.cd_contato_ddd1.value.length==""){window.alert ("Informe o ddd.");shipinfo.cd_contato_ddd1.focus();return (false);}
	if (shipinfo.nm_contato_numero1.value.length==""){window.alert ("Informe o telefone.");shipinfo.nm_contato_numero1.focus();return (false);}
	if (shipinfo.cd_contato_email3.value.length==""){window.alert ("Informe o e-mail.");shipinfo.cd_contato_email3.focus();return (false);}
	if (shipinfo.nm_contato_nome4.value.length==""){window.alert ("Informe o nome do contato para urgências.");shipinfo.nm_contato_nome4.focus();return (false);}
	if (shipinfo.nm_contato_cargo4.value.length==""){window.alert ("Informe o tipo de relacionamento do contato.");shipinfo.nm_contato_cargo4.focus();return (false);}
	if (shipinfo.cd_contato_ddd4.value.length==""){window.alert ("Informe o telefone de urgência.");shipinfo.cd_contato_ddd4.focus();return (false);}
	if (shipinfo.nm_contato_numero4.value.length==""){window.alert ("Informe o telefone de urgência.");shipinfo.nm_contato_numero4.focus();return (false);}	
	if (shipinfo.cd_pis.value.length==""){window.alert ("Informe o nº do PIS.");shipinfo.cd_pis.focus();return (false);}
	if (shipinfo.cd_ctps.value.length==""){window.alert ("Informe o nº da carteira de trabalho.");shipinfo.cd_ctps.focus();return (false);}
	if (shipinfo.cd_ctps_serie.value.length==""){window.alert ("Informe a série da carteira de trabalho.");shipinfo.cd_ctps_serie.focus();return (false);}
	if (shipinfo.nm_rg.value.length==""){window.alert ("Informe o RG.");shipinfo.nm_rg.focus();return (false);}
	if (shipinfo.dt_rg_emissao.value.length==""){window.alert ("Informe a data de emissão do RG.");shipinfo.dt_rg_emissao.focus();return (false);}
	if (shipinfo.nm_rg_expedidor.value.length==""){window.alert ("Informe o Orgão expedidor do RG.");shipinfo.nm_rg_expedidor.focus();return (false);}
	if (shipinfo.nm_cpf.value.length==""){window.alert ("Informe o CPF.");shipinfo.nm_cpf.focus();return (false);}
	if (shipinfo.nm_tit_eleitor.value.length==""){window.alert ("Informe o nº do Título Eleitoral.");shipinfo.nm_tit_eleitor.focus();return (false);}
	if (shipinfo.nr_tit_eleitor_zona.value.length==""){window.alert ("Informe a Zona Eleitoral.");shipinfo.nr_tit_eleitor_zona.focus();return (false);}
	if (shipinfo.nr_tit_eleitor_seccao.value.length==""){window.alert ("Informe a Secção Eleitoral.");shipinfo.nr_tit_eleitor_seccao.focus();return (false);}
		
	if (shipinfo.nr_caract_altura.value.length==""){window.alert ("Informe a altura.");shipinfo.nr_caract_altura.focus();return (false);}
	if (shipinfo.nr_caract_peso.value.length==""){window.alert ("Informe o peso.");shipinfo.nr_caract_peso.focus();return (false);}
	if (shipinfo.nr_caract_manequim.value.length==""){window.alert ("Informe o nº do manequim.");shipinfo.nr_caract_manequim.focus();return (false);}
	if (shipinfo.nr_caract_calcado.value.length==""){window.alert ("Informe o nº do calçado");shipinfo.nr_caract_calcado.focus();return (false);}
	
	if (shipinfo.cd_sexo.value==1){
			if (shipinfo.nm_reservista.value.length==""){window.alert ("Informe o nº da Reservista.");shipinfo.nm_reservista.focus();return (false);}		
		}
	
	if (shipinfo.cd_estado_civil.value==2){
			if (shipinfo.nm_conjuge.value.length==""){window.alert ("Informe o nome do(a) conjugê.");shipinfo.nm_conjuge.focus();return (false);}
			if (shipinfo.dt_conjuge_nasc.value.length==""){window.alert ("Informe a data de nascimento do(a) conjugê.");shipinfo.dt_conjuge_nasc.focus();return (false);}
		}
		
	if (shipinfo.cd_conspro.value==2){
			if (shipinfo.cd_numreg.value.length==""){window.alert ("Informe o nº do COREN.");shipinfo.cd_numreg.focus();return (false);}
			if (shipinfo.rgpro_concessao.value.length==""){window.alert ("Informe o tipo de concessão.");shipinfo.rgpro_concessao.focus();return (false);}
			if (shipinfo.rgpro_status.value.length==""){window.alert ("Informe a situação.");shipinfo.rgpro_status.focus();return (false);}
			if (shipinfo.dt_rgproinscr.value.length==""){window.alert ("Informe a data de inscrição.");shipinfo.dt_rgproinscr.focus();return (false);}
			if (shipinfo.dt_rgpropag.value.length==""){window.alert ("Informe a data de pagamento.");shipinfo.dt_rgpropag.focus();return (false);}
			if (shipinfo.dt_rgproval.value.length==""){window.alert ("Informe a data de validade.");shipinfo.dt_rgproval.focus();return (false);}
		}		
	return (true);}
	
</script>
<!--#include file="../../js/ajax.js"-->
<script language="Javascript">
<!-- 
//-------------------------------------------------------------------
function rep_space(string) {
	while (string.indexOf('a') != -1) {
 		string = string.replace('a','%20');
	}
	return string;
}
//--------------------------------------------------------------------


//---------------------------------------------------------------------->
</script>
<style type="text/css" media="screen">
<!--
.ok_print{ visibility:hidden; display: none;}
.no_print{ visibility:visible; display:block;}
#frame{border:0px inset;
	position:absolute;
	height:93%;
	width:79%;
	top:20px;
	left:210px;
	margin: 0px;
	padding: 6px;
	text-align:left;
	overflow:scroll;
	text-decoration:none;}
table{background-color: #ffffff; 
    border: 1px solid #cccccc;
	font-family: verdana;
	font-size: 9px;
	text-decoration:none;}

a:link {color:#000000;text-decoration:none;}
a:visited {color:#000000;text-decoration:none;}
a:hover {color:#FF0000;}
a:active {color:#000000;}

.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; } 

-->
</style>

<%cor_campo = "#ebebeb"%>
					
					<%If strmsg <> "" then%>
						<table  border="0" cellspacing="0" cellpadding="0" class="txt_cinza">
							<tr class="no_print">
							 <td>
								<b><%=strmsg%></b>
							 </td>
							</tr>
							<tr class="no_print">
							 <td align=center><a href="coop_cooperados_lista.asp"><img src="../../imagens/bt_lista.gif" alt="" width="119" height="22" border="1"></a></td>
							</tr>
						</table>
					<%else%>
					<table border="1" cellspacing="0" cellpadding="2" align="center" class="nok_print" style="">
						<form name="form" action="../acoes/funcionarios_inicial_acao.asp"  method="POST" onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="cd_sessao" value="<%=session_cd_usuario%>">
						<input type="hidden" name="jan" value="1">
						<input type="hidden" name="hr_entrada" value="08:00">
						<input type="hidden" name="hr_saida" value="18:00">
						<input type="hidden" name="nm_intervalo" value="">
						<input type="hidden" name="cd_unidade" value="27">
						<input type="hidden" name="cd_status" value="1">
						<input type="hidden" name="cd_cargo" value="42">
						<input type="hidden" name="cd_regime_trabalho" value="">
						<input type="hidden" name="cd_dia" value="">
						<input type="hidden" name="cd_mes" value="">
						<input type="hidden" name="cd_ano" value="">
						<%If strcod = "" then%>
						<input type="hidden" name="acao" value="insert">
						<%ElseIf strcod <> ""  then %>
						<input type="hidden" name="acao" value="altera">
						<input type="hidden" name="cod" value="<%=strcod%>">
						<input type="hidden" name="cd_funcionario" value="<%=str_funcionario%>">
						<input name="foto_h" type="hidden" value="<%=strfoto%>">
						
						<%End if%>
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><b>FICHA CADASTRAL INICIAL<%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/px.gif" width="712" height="1">&nbsp;</td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="3"><b>Nome</b><br><input type="text" name="nm_nome" value="<%=strnome %>" size="65" maxlength="70" class="inputs"></td>
								<td>
									<b>Sexo</b><br>
									<select name="cd_sexo"  class="inputs">
											<option value=""></option>
											<option value="1" <%'if str_sexo = 1 then response.write("SELECTED")%>>Masculino</option>
											<option value="2" <%'if str_sexo = 2 then response.write("SELECTED")%>>Feminino</option>
									</select>								
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><b>Endereço / Nº</b><br>
								<input type="text" name="nm_endereco" value="<%=strnm_endereco%>" size="65" maxlength="100" class="inputs">&nbsp;
								<input type="text" name="nr_numero" value="<%=strnr_numero%>" size="4" maxlength="8" class="inputs"></td>
								
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="1"><b>Complemento</b><br><input type="text" name="nm_complemento" value="<%=strnm_complemento%>" size="25" maxlength="30" class="inputs"></td>
								<td><b>Bairro</b><br><input type="text" name="nm_bairro" value="<%=strbairro%>" size="20" maxlength="100" class="inputs"></td>
								<td><b>CEP</b><br><input type="text" name="nm_cep" value="<%=strnm_cep%>" size="10" maxlength="9" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2"><b>Cidade</b><br><input type="text" name="nm_cidade" value="<%=strnm_cidade%>" size="50" maxlength="30" class="inputs"></td>
								<td><b>UF</b><br><input type="text" name="nm_estado" value="<%=strnm_estado%>" size="1" maxlength="2" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Data de nascimento</b><br>
									<input type="text" name="dia_nascimento" value="<%=strdia_nascimento%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
									<input type="text" name="mes_nascimento" value="<%=strmes_nascimento%>" size="2" maxlength="2" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
									<input type="text" name="ano_nascimento" value="<%=strano_nascimento%>" size="4" maxlength="4" class="inputs">
								<td><b>Nacionalidade (País)</b><br><input type="text" name="nm_nacionalidade" value="<%=str_nacionalidade%>" size="25" maxlength="25" class="inputs"></td>
								<td><b>Naturalidade (Cidade)</b><br><input type="text" name="nm_naturalidade" value="<%=str_naturalidade%>" size="25" maxlength="30" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">								
								<td><b>Estado Civil</b><br>
									<select name="cd_estado_civil" class="inputs">
									<option value=""></option>
									<%strsql = "Select * From TBL_estado_civil order by cd_codigo"
									Set rs_est_civil = dbconn.execute(strsql)%>
									<%Do While Not rs_est_civil.eof
											cd_estado_civil = rs_est_civil("cd_codigo")
											nm_estado_civil = rs_est_civil("nm_estado_civil")%>	
											<option value="<%=cd_estado_civil%>" <%if int(str_estado_civil) = cd_estado_civil then%>Selected<%end if%>><%=nm_estado_civil%></option>
											<%rs_est_civil.movenext
										loop
											rs_est_civil.close
											Set rs_est_civil = nothing%>											
									</select>
								</td><td colspan="2"><b>Nome do Cônjuge</b> <img src="../../imagens/px.gif" width="145" height="1"> <b>Nasc.</b><br><input type="text" name="nm_conjuge" value="<%=str_conjuge%>" size="33" maxlength="100" class="inputs"> &nbsp;  &nbsp;  &nbsp;
								<input type="text" name="dia_conjuge_nasc" value="" size="2" maxlength="10" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
								<input type="text" name="mes_conjuge_nasc" value="" size="2" maxlength="10" class="inputs" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
								<input type="text" name="ano_conjuge_nasc" value="" size="4" maxlength="10" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<!--tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><b>Endereço alternativo</b><br><input type="text" name="nm_endereco_alt" value="<%=str_endereco_alt%>" size="100 maxlength="150" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4">&nbsp;</td></tr-->
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/px.gif" width="100%" height="1"></td></tr>
<%'cor_campo = "#ffffff"%>	
							<tr bgcolor="#a1a1a1"><td colspan="4" style="color:white;"><b>FILIAÇÃO</b></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><b>Mãe</b><img src="../imagens/px.gif" width="63" height="1"><input type="text" name="nm_mae" value="<%=strmae%>" size="86" maxlength="100" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><b>Pai</b><img src="../imagens/px.gif" width="67" height="1"><input type="text" name="nm_pai" value="<%=strpai%>" size="86" maxlength="100" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ffffff"%>	
							<tr bgcolor="#a1a1a1"><td colspan="4" style="color:white;"><b>CONTATOS</b></td></tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<table border="0" style="border-collapse:collapse; border:solid white 0px; ">
										<tr bgcolor="<%=cor_campo%>">
											<!--td><b>Nome</b></td-->
											<td align="left"><b>Tipo</b><br><img src="../../imagens/px.gif" width="60" height="1"></td>
											<td align="left"><b>Tel./Cel/E-mail/Nextel</b><br><img src="../../imagens/px.gif" width="190" height="1"></td>
											<td align="left"><b>Observações</b><br><img src="../../imagens/px.gif" width="200" height="1"></td>
											<td>&nbsp;</td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<input type="hidden" name="cd_contato_origem" value="4">
											<input type="hidden" name="id_contato_origem" value="<%=str_funcionario%>">
											<td>Tel/cel: <input type="hidden" name="cd_contato_categoria1" value="1"></td>
											<td><input type="text" name="cd_contato_ddd1" value="<%=str_contato_ddd%>" size="2" maxlength="2" class="inputs">&nbsp;
												<input type="text" name="nm_contato_numero1" value="<%=str_contato%>" size="20" maxlength="100" class="inputs"></td>
											<td><input type="text" name="nm_contato_obs1" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs"></td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<td>Tel/cel: <input type="hidden" name="cd_contato_categoria2" value="2"></td>
											<td><input type="text" name="cd_contato_ddd2" value="<%=str_contato_ddd%>" size="2" maxlength="2" class="inputs">&nbsp;
												<input type="text" name="nm_contato_numero2" value="<%=str_contato%>" size="20" maxlength="100" class="inputs"></td>
											<td><input type="text" name="nm_contato_obs2" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs"></td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<td>E-mail: <input type="hidden" name="cd_contato_categoria3" value="3"></td>
											<td><input type="text" name="cd_contato_email3" value="<%=str_contato%>" size="23" maxlength="100" class="inputs"></td>
											<td><input type="text" name="nm_contato_obs3" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs"></td>
										</tr>
										<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
										<tr bgcolor="<%=cor_campo%>">
											<!--td><b>Nome</b></td-->
											<td align="left"><b>Tipo</b><br><img src="../../imagens/px.gif" width="60" height="1"></td>
											<td align="left"><b>Nome<img src="../../imagens/px.gif" width="50" height="1"> Grau relac.</b><br><img src="../../imagens/px.gif" width="180" height="1"></td>
											<td align="left"><b>Tel./Cel/E-mail/Nextel</b><br><img src="../../imagens/px.gif" width="190" height="1"></td>
											<td align="left"><b>Observações</b><br><img src="../../imagens/px.gif" width="200" height="1"></td>
											<td>&nbsp;</td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<td style="color:red;">Urgência: <input type="hidden" name="cd_contato_categoria4" value="5"></td>
											<td><input type="text" name="nm_contato_nome4" value="<%=str_contato_nome%>" size="10" maxlength="50" class="inputs">&nbsp;
												<input type="text" name="nm_contato_cargo4" value="<%=str_contato_cargo%>" size="10" maxlength="50" class="inputs">&nbsp;</td>
											<td><input type="text" name="cd_contato_ddd4" value="<%=str_contato_ddd%>" size="2" maxlength="2" class="inputs">&nbsp;
												<input type="text" name="nm_contato_numero4" value="<%=str_contato%>" size="20" maxlength="100" class="inputs"></td>
											<td><input type="text" name="nm_contato_obs4" value="<%=str_contato_obs%>" size="25" maxlength="100" class="inputs"></td>
											
										</tr>
									</table>
								</td>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1">
								<td colspan="4" style="color:white;"><b>DOCUMENTAÇÃO</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">								
								<td><b>PIS</b><br><input type="text" name="cd_pis" value="<%=str_pis%>" size="20" maxlength="25" class="inputs"></td>
								<td><b>Carteira Profissional</b><br><input type="text" name="cd_ctps" value="<%=str_ctps%>" size="20" maxlength="20" class="inputs"></td>
								<td><b>Série</b><br><input type="text" name="cd_ctps_serie" size="5" class="inputs" maxlength="5" value="<%=str_ctps_serie%>"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>RG</b><br><input type="text" name="nm_rg" value="<%=strnm_rg%>" size="20" maxlength="15" class="inputs"></td>
								<td><b>Data emissão &nbsp; - &nbsp; Órgão exped.</b><br><input type="text" name="dt_rg_emissao" value="<%=strdt_rg_emissao%>" size="10" maxlength="15" class="inputs"> &nbsp;
									<input type="text" name="nm_rg_expedidor" value="<%=strnm_rg_expedidor%>" size="6" maxlength="50" class="inputs"></td>
								<td><b>CPF</b><br><input type="text" name="nm_cpf" value="<%=strnm_cpf%>" size="25" maxlength="14" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">								
								<td><b>Tit. Eleitor</b><br><input type="text" name="nm_tit_eleitor" value="<%=strnm_tit_eleitor%>" size="25" maxlength="14" class="inputs"></td>
								<td><b>Zona &nbsp; &nbsp; &nbsp; - &nbsp; Secção</b><br><input type="text" name="nr_tit_eleitor_zona" value="<%=strnr_tit_eleitor_zona%>" size="5" maxlength="14" class="inputs"> &nbsp;
									<input type="text" name="nr_tit_eleitor_seccao" value="<%=strnr_tit_eleitor_seccao%>" size="5" maxlength="14" class="inputs"></td>
								<td><b>Reservista</b><br><input type="text" name="nm_reservista" value="<%=strnm_reservista%>" size="25" maxlength="14" class="inputs"></td>
								<td>&nbsp;</td>
							</tr>
							
							<!--*****************-->
							<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr-->
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><a name="coren"></a><b>Registro Profissional</b><br>
									<select name="cd_conspro" class="inputs">
										<option value="1"></option>
										<%strsql = "Select * From TBL_cons_profissional "
					  					Set rs_conspro = dbconn.execute(strsql)
											While not rs_conspro.EOF
												cod_conspro = rs_conspro("cd_codigo")
												nm_conspro = rs_conspro("nm_conspro")
													if cod_conspro = int(str_conspro) then
														conspro_sel = "Selected"
													else
														conspro_sel = ""
													end if%>
																		
												<option value="<%=cod_conspro%>" <%=conspro_sel%>><%=nm_conspro%></option>
												<%rs_conspro.movenext
											wend%>
									</select></td>
								<td><b>Número</b><br><input type="text" name="cd_numreg" value="<%=str_numreg%>" size="20" maxlength="25" class="inputs"></td>
								<td valign="top"><b>Tipo de concessão / Situação</b><br>
									<select name="rgpro_concessao" class="inputs">
											<%if str_rgpro_concessao = "" then
												concessao_sel_0 =	"SELECTED"
											elseif str_rgpro_concessao = 1 then
												concessao_sel_1 =	"SELECTED"
											elseif str_rgpro_concessao = 2 then
												concessao_sel_2 =	"SELECTED"
											end if%>
											<option value="" <%=concessao_sel_0%>></option>
											<option value="1" <%=concessao_sel_1%>>Definitivo</option>
											<option value="2" <%=concessao_sel_2%>>Provisório</option>
									</select>&nbsp;									
									<select name="rgpro_status" class="inputs">
											<%if str_rgpro_status = "" then
												rgpro_status_sel_0 =	"SELECTED"
											elseif str_rgpro_status = 1 then
												rgpro_status_sel_1 =	"SELECTED"
											elseif str_rgpro_status = 2 then
												rgpro_status_sel_2 =	"SELECTED"
											end if%>
											<option value="" <%=rgpro_status_sel_0%>></option>
											<option value="1" <%=rgpro_status_sel_1%>>Ativo</option>
											<option value="2" <%=rgpro_status_sel_2%>>Inativo</option>
									</select></td>
								<td valign="top" rowspan="3"><b>Observações</b><br><textarea cols="25" rows="4" name="obs_rgpro" class="inputs"><%=str_rgpro%></textarea></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td valign="top"><b>Data de inscrição</b><br><input type="text" name="dt_rgproinscr" size="20" maxlength="10" class="inputs" value="<%=str_dt_rgproinscr%>"></td>
								<td valign="top"><b>Data de pagamento</b><br><input type="text" name="dt_rgpropag" size="20" maxlength="10" class="inputs" value="<%=str_dt_rgpropag%>"></td>
								<td valign="top"><b>Validade</b><br><input type="text" name="dt_rgproval" size="25" maxlength="12" class="inputs" value="<%=str_dt_rgproval%>"></td>
								<!--td><input type="text" name="dia_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="mes_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="ano_valid" size="5" maxlength="4" class="inputs"></td-->
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1">
								<td colspan="4" style="color:white;"><b>CARTEIRA DE VACINAÇÃO</b><a name="vacinas"></a></td>
							</tr>

							<tr bgcolor="<%=cor_campo%>">
								<td><b>Hepatite B</b></td>
								<td colspan="3">
									<%if int(strnr_hepatite_b) = 3 then
										ckhepatite_b3 = "checked"
										ckhepatite_b2 = "checked"
										ckhepatite_b1 = "checked"
									elseif int(strnr_hepatite_b) = 2 then
										ckhepatite_b2 = "checked"
										ckhepatite_b1 = "checked"
									elseif int(strnr_hepatite_b) = 1 then
										ckhepatite_b1 = "checked"
									end if%>
									<input type="radio" name="nr_hepatite_b" value="1" style="border:0;" class="inputs" <%=ckhepatite_b0%>>Nenhuma dose &nbsp; 
									<input type="radio" name="nr_hepatite_b" value="1" style="border:0;" class="inputs" <%=ckhepatite_b1%>>1 dose &nbsp; 
									<input type="radio" name="nr_hepatite_b" value="2" style="border:0;" class="inputs" <%=ckhepatite_b2%>>2 doses &nbsp;
									<input type="radio" name="nr_hepatite_b" value="3" style="border:0;" class="inputs" <%=ckhepatite_b3%>>3 doses</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Dupla Adulto</b></td>
								<td colspan="3">
									<%if int(strnr_dupla_adulto) = 3 then
										ckdp_adulto_b3 = "checked"
										ckdp_adulto_b2 = "checked"
										ckdp_adulto_b1 = "checked"
									elseif int(strnr_dupla_adulto) = 2 then
										ckdp_adulto_b2 = "checked"
										ckdp_adulto_b1 = "checked"
									elseif int(strnr_dupla_adulto) = 1 then
										ckdp_adulto_b1 = "checked"
									end if%>
									<input type="radio" name="nr_dupla_adulto" value="0" style="border:0;" class="inputs" <%=ckdp_adulto_b0%>>Nenhuma dose &nbsp; 
									<input type="radio" name="nr_dupla_adulto" value="1" style="border:0;" class="inputs" <%=ckdp_adulto_b1%>>1 dose &nbsp; 
									<input type="radio" name="nr_dupla_adulto" value="2" style="border:0;" class="inputs" <%=ckdp_adulto_b2%>>2 doses &nbsp;
									<input type="radio" name="nr_dupla_adulto" value="3" style="border:0;" class="inputs" <%=ckdp_adulto_b3%>>3 doses &nbsp; &nbsp; 
									<b>Val.:</b> <input type="text" name="dt_dupla_adulto_validade" value="<%=strdt_dupla_adulto_validade%>" size="10" maxlength="10" class="inputs"></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>SCR (Sarampo, Caxumba, Rubeóla)</b></td>
								<td><input type="radio" name="nr_scr" value="1" <%if int(strnr_scr) = 1 then%>checked<%end if%> style="border:0;" class="inputs"> SIM &nbsp; &nbsp; &nbsp; 
									<input type="radio" name="nr_scr" value="" <%if int(strnr_scr) = "" OR int(strnr_scr) = 0 then%>checked<%end if%> style="border:0;" class="inputs"> Não</td>								
								<td colspan="2">&nbsp;</td>
							</tr>
							<!--tr bgcolor="<%=cor_campo%>"><td colspan="4"><a name="exammed"><br><img src="../../imagens/blackdot.gif" width="100%" height="1"></a></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="2"><b>Exame médico realizado em: </b><input type="text" name="dt_exame_medico" value="<%=strdt_exame_medico%>" size="10" maxlength="10" class="inputs"></td>
								<td colspan="2"><b>Validade: </b><%if strdt_exame_medico <> "" Then%><%=day(strdt_exame_medico)&"/"&month(strdt_exame_medico)&"/"&year(strdt_exame_medico)+1%><%end if%></td>
							</tr-->
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1">
								<td style="color:white;" colspan="4"><b>CARACTERÍSTICAS FÍSICAS</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Altura</b><br><input type="text" name="nr_caract_altura" value="<%=strnr_caract_altura%>" size="6" maxlength="6" class="inputs"></td>
								<td><b>Peso </b><br><input type="text" name="nr_caract_peso" value="<%=strnr_caract_peso%>" size="6" maxlength="6" class="inputs"></td>
								<td><b>Manequim</b><br><input type="text" name="nr_caract_manequim" value="<%=strnr_caract_manequim%>" size="6" maxlength="6" class="inputs"></td>
								<td><b>Calçado</b><br><input type="text" name="nr_caract_calcado" value="<%=strnr_caract_calcado%>" size="6" maxlength="6" class="inputs"></td>
							</tr>
							
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1">
								<td colspan="4" style="color:white;"><b>DEPENDENTES (FILHOS)</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
								
									<table style="border:solid white 0px; border-collapse:collapse;">
										<tr bgcolor="<%=cor_campo%>">
											<td><b>Nome</b><br><img src="../../imagens/px.gif" alt="" width="160" height="1" border="0"></td>
											<td><b>Sexo</b><br><img src="../../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
											<td><b>Data nasc.</b><br><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
											<td><b>Observações.</b><br><img src="../../imagens/px.gif" alt="" width="110" height="1" border="0"></td>
											<td>&nbsp;</td>
										</tr>
										<!--form name="formdependente"-->
										<tr bgcolor="<%=cor_campo%>">
											<td><input type="text" name="nm_dependente_nome1" size="40" maxlength="30" class="inputs"></td>
												<td><select name="cd_dependente_sexo1" class="inputs">
													<option value=""></option>
													<option value="1">Masc.</option>
													<option value="2">Fem.</option>
												</select>
												</td>
											<td><input type="text" name="dt_dependente_nascimento1" value="" size="10" maxlength="10" class="inputs"></td>
											<td><input type="text" name="cd_dependente_obs1" value="" size="15" class="inputs"></td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<td><input type="text" name="nm_dependente_nome2" size="40" maxlength="30" class="inputs"></td>
											<td><select name="cd_dependente_sexo2" class="inputs">
													<option value=""></option>
													<option value="1">Masc.</option>
													<option value="2">Fem.</option>
												</select>
												</td>
											<td><input type="text" name="dt_dependente_nascimento2" value="" size="10" maxlength="10" class="inputs"></td>
											<td><input type="text" name="cd_dependente_obs2" value="" size="15" class="inputs"></td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<td><input type="text" name="nm_dependente_nome3" size="40" maxlength="30" class="inputs"></td>
											<td><select name="cd_dependente_sexo3" class="inputs">
													<option value=""></option>
													<option value="1">Masc.</option>
													<option value="2">Fem.</option>
												</select>
												</td>
											<td><input type="text" name="dt_dependente_nascimento3" value="" size="10" maxlength="10" class="inputs"></td>
											<td><input type="text" name="cd_dependente_obs3" value="" size="15" class="inputs"></td>
										</tr>
										<tr bgcolor="<%=cor_campo%>">
											<td><input type="text" name="nm_dependente_nome4" size="40" maxlength="30" class="inputs"></td>
											<td><select name="cd_dependente_sexo4" class="inputs">
													<option value=""></option>
													<option value="1">Masc.</option>
													<option value="2">Fem.</option>
												</select>
												</td>
											<td><input type="text" name="dt_dependente_nascimento4" value="" size="10" maxlength="10" class="inputs"></td>
											<td><input type="text" name="cd_dependente_obs4" value="" size="15" class="inputs"></td>
										</tr>
									</table>
									
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>ESCOLARIDADE / FORMAÇÃO</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4"><b>Escolaridade</b><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0">
								<b>Instituição</b><img src="../../imagens/px.gif" alt="" width="85" height="1" border="0">
								<b>Curso</b><img src="../../imagens/px.gif" alt="" width="105" height="1" border="0"> 
								<b>andamento</b><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"> 
								<b>Término</b><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0">
								<b>Observação</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">								
									<select name="cd_ensino_grau1" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_escolaridade_grau order by cd_codigo"
										Set rs_grau = dbconn.execute(xsql)
										while not rs_grau.EOF
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<option value="<%=cd_grau%>"><%=nm_grau%></option>
										<%rs_grau.movenext
										Wend%>
									</select> &nbsp; 
								<input type="text" name="nm_instituicao1" class="inputs" size="17">
								<input type="text" name="nm_curso1" class="inputs" size="17">
									<select name="cd_ensino_andamento1" class="inputs">
										<option value="0"></option>
										<%xsql = "Select * From TBL_escolaridade_andamento order by cd_codigo DESC"
										Set rs_andamento = dbconn.execute(xsql)
										while not rs_andamento.EOF
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<option value="<%=cd_andamento%>"><%=nm_ensino_andamento%></option>
										<%rs_andamento.movenext
										Wend%>
									</select>
									<input type="text" name="dt_termino1" maxlength="10" size="10" class="inputs">
									<input type="text" name="nm_obs1" size="20" maxlength="200" class="inputs">
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">								
									<select name="cd_ensino_grau2" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_escolaridade_grau order by cd_codigo"
										Set rs_grau = dbconn.execute(xsql)
										while not rs_grau.EOF
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<option value="<%=cd_grau%>"><%=nm_grau%></option>
										<%rs_grau.movenext
										Wend%>
									</select> &nbsp; 
								<input type="text" name="nm_instituicao2" class="inputs" size="17">
								<input type="text" name="nm_curso2" class="inputs" size="17">
									<select name="cd_ensino_andamento2" class="inputs">
										<option value="0"></option>
										<%xsql = "Select * From TBL_escolaridade_andamento order by cd_codigo DESC"
										Set rs_andamento = dbconn.execute(xsql)
										while not rs_andamento.EOF
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<option value="<%=cd_andamento%>"><%=nm_ensino_andamento%></option>
										<%rs_andamento.movenext
										Wend%>
									</select>
									<input type="text" name="dt_termino2" maxlength="10" size="10" class="inputs">
									<input type="text" name="nm_obs2" size="20" maxlength="200" class="inputs">
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">								
									<select name="cd_ensino_grau3" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_escolaridade_grau order by cd_codigo"
										Set rs_grau = dbconn.execute(xsql)
										while not rs_grau.EOF
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<option value="<%=cd_grau%>"><%=nm_grau%></option>
										<%rs_grau.movenext
										Wend%>
									</select> &nbsp; 
								<input type="text" name="nm_instituicao3" class="inputs" size="17">
								<input type="text" name="nm_curso3" class="inputs" size="17">
									<select name="cd_ensino_andamento3" class="inputs">
										<option value="0"></option>
										<%xsql = "Select * From TBL_escolaridade_andamento order by cd_codigo DESC"
										Set rs_andamento = dbconn.execute(xsql)
										while not rs_andamento.EOF
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<option value="<%=cd_andamento%>"><%=nm_ensino_andamento%></option>
										<%rs_andamento.movenext
										Wend%>
									</select>
									<input type="text" name="dt_termino3" maxlength="10" size="10" class="inputs">
									<input type="text" name="nm_obs3" size="20" maxlength="200" class="inputs">
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">								
									<select name="cd_ensino_grau4" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_escolaridade_grau order by cd_codigo"
										Set rs_grau = dbconn.execute(xsql)
										while not rs_grau.EOF
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<option value="<%=cd_grau%>"><%=nm_grau%></option>
										<%rs_grau.movenext
										Wend%>
									</select> &nbsp; 
								<input type="text" name="nm_instituicao4" class="inputs" size="17">
								<input type="text" name="nm_curso4" class="inputs" size="17">
									<select name="cd_ensino_andamento4" class="inputs">
										<option value="0"></option>
										<%xsql = "Select * From TBL_escolaridade_andamento order by cd_codigo DESC"
										Set rs_andamento = dbconn.execute(xsql)
										while not rs_andamento.EOF
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<option value="<%=cd_andamento%>"><%=nm_ensino_andamento%></option>
										<%rs_andamento.movenext
										Wend%>
									</select>
									<input type="text" name="dt_termino4" maxlength="10" size="10" class="inputs">
									<input type="text" name="nm_obs4" size="20" maxlength="200" class="inputs">
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ffffff"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>EXPERIÊNCIA PROFISSIONAL</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<b>Emprego Anterior</b><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0">
									<b>Função</b><img src="../../imagens/px.gif" alt="" width="130" height="1" border="0">
									<b>Admissão</b><img src="../../imagens/px.gif" alt="" width="50" height="1" border="0">
									<b>Desligamento</b><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0">
									<b>Obs.</b></td>
							</tr>
							<!--tr><td colspan="4"--><input type="hidden" name="emprego_total" value="" size="100">
								<input type="hidden" name="emprego_result" size="70"><!--/td></tr-->
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<input type="text" name="nm_emprego_empresa1" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="nm_emprego_cargo1" class="inputs"><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0">
									<input type="text" name="dt_emprego_entrada1" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_saida1" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_obs1" size="20" maxlength="200" class="inputs">									
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<input type="text" name="nm_emprego_empresa2" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="nm_emprego_cargo2" class="inputs"><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0">
									<input type="text" name="dt_emprego_entrada2" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_saida2" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_obs2" size="20" maxlength="200" class="inputs">									
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<input type="text" name="nm_emprego_empresa3" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="nm_emprego_cargo3" class="inputs"><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0">
									<input type="text" name="dt_emprego_entrada3" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_saida3" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_obs3" size="20" maxlength="200" class="inputs">									
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td colspan="4">
									<input type="text" name="nm_emprego_empresa4" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="nm_emprego_cargo4" class="inputs"><img src="../../imagens/px.gif" alt="" width="30" height="1" border="0">
									<input type="text" name="dt_emprego_entrada4" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_saida4" maxlength="10" size="10" class="inputs"><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0">
									<input type="text" name="dt_emprego_obs4" size="20" maxlength="200" class="inputs">									
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
<%cor_campo = "#ebebeb"%>
							<tr bgcolor="#a1a1a1" >
								<td colspan="4" style="color:white;"><b>INFORMAÇÕES COMPLEMENTARES</b></td>
							</tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Possui parente na Empresa? </b><!-- <%'if strnm_parente <> "" Then%>SIM<%'Else%>NÃO<%'end if%>--></td>
								<td colspan="2"><b>Nome: </b><input type="text" name="nm_parente" value="<%=strnm_parente%>" size="50" maxlength="50" class="inputs"></td>
								<td><b>Parentesco: </b>
									<select name="cd_parente_parentesco" class="inputs">
										<option value=""></option>
										<%xsql = "Select * From TBL_parentesco order by nm_parentesco DESC"
										Set rs_parentesco = dbconn.execute(xsql)
										while not rs_parentesco.EOF
										cd_parentesco = rs_parentesco("cd_codigo")
										nm_parentesco = rs_parentesco("nm_parentesco")%>
											<option value="<%=cd_parentesco%>" <%if int(strcd_parente_parentesco) = cd_parentesco then%> SELECTED<%end if%>><%=nm_parentesco%></option>
										<%rs_parentesco.movenext
										Wend%>
									</select>
								</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Residência: (tipo)</b><br><select name="cd_residencia_tipo" class="inputs">
										<option value="0"></option>
										<option value="1" <%if int(strcd_residencia_tipo) = 1 Then%>selected<%end if%>><b>Própria<b></option>
										<option value="2" <%if int(strcd_residencia_tipo) = 2 Then%>selected<%end if%>><b>Alugada<b></option>
									</select>
								<td><b>Financiada?</b><br><select name="cd_residencia_financ" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_residencia_financ) = 1 Then%>selected<%end if%>>Sim</option>
										<option value="2" <%if int(strcd_residencia_financ) = 2 Then%>selected<%end if%>>Não</option>
									</select></td>
								<td><b>Reside</b><br><select name="cd_residencia_reside" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_residencia_reside) = 1 Then%>selected<%end if%>>Só</option>
										<option value="2" <%if int(strcd_residencia_reside) = 2 Then%>selected<%end if%>>Companheiro</option>
										<option value="3" <%if int(strcd_residencia_reside) = 3 Then%>selected<%end if%>>Pais</option>
										<option value="4" <%if int(strcd_residencia_reside) = 4 Then%>selected<%end if%>>Parentes</option>
										<option value="5" <%if int(strcd_residencia_reside) = 5 Then%>selected<%end if%>>Amigos</option>
									</select></td>
								<td>&nbsp;</td>
							</tr>
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
							<tr bgcolor="<%=cor_campo%>">
								<td><b>Veículo: (tipo)</b><br><select name="cd_veiculo_tipo" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_veiculo_tipo) = 1 Then%>selected<%end if%>>Carro</option>
										<option value="2" <%if int(strcd_veiculo_tipo) = 2 Then%>selected<%end if%>>Moto</option>
									</select>				
								</td>
								<td><b>Financiado?</b><br><select name="cd_veiculo_financ" class="inputs">
										<option value=""></option>
										<option value="1" <%if int(strcd_veiculo_financ) = 1 Then%>selected<%end if%>>Sim</option>
										<option value="2" <%if int(strcd_veiculo_financ) = 2 Then%>selected<%end if%>>Não</option>
									</select></td>
								<td colspan="2">&nbsp;</td>
							</tr>
							<input type="hidden" name="valores_total" value="" size="100">
							<input type="hidden" name="valores_result" size="70">
							
							
							<tr bgcolor="<%=cor_campo%>"><td colspan="4"><img src="../../imagens/blackdot.gif" width="100%" height="1"></td></tr>
<%cor_campo = "#ebebeb"%>
							
							
							<tr class="no_print">
								<td colspan="4" valign="middle"><!--<A href="empresa.asp?tipo=lista"><img src="../../imagens/bt_lista.gif" alt="Listar" border="0" width="119" height="22" border="1"></a>-->
								<input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0">
								<%=stracao%></td>
							</tr>
							
							<%'end if%>
							</form>
							<%If strcod <> "" then%>							
								<tr>
									<td>
									<form name="form_ex" action="empresa/acoes/empr_funcionarios_acao.asp" id="forms" method="post"  enctype="multipart/form-data">
										<input type="hidden" name="acao" value="excluir">
										<input type="hidden" name="cod" value="<%=cd_funcionario%>">
									</form></td>
								</tr>
							<%End if%>
						</table>
									 								
					<%End if%>
			<SCRIPT LANGUAGE="javascript">
 {
        //MaskInput(document.forms.nm_cep, "99999-99");
		//MaskInput(document.forms.dt_nascimento, "99/99/9999");	    
 }
</SCRIPT>
			
			<SCRIPT LANGUAGE="javascript">
				function JsDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o funcionario?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir&protege_exclusao='+cod2+'');
					}
					  }
				
				function JsContatoDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o contato?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_cont='+cod2+'&acao=apaga_contato');
					}
					  }  
					  
				function JsDependenteDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir o dependente?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_dep='+cod2+'&acao=apaga_dependente');
					}
					  }  
				
				function JsEscolaridadeDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de escolaridade?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_esc='+cod2+'&acao=apaga_escolaridade');
					}
					  }    
				
				function JsEmpregodeDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse registro de emprego anterior?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_emp='+cod2+'&acao=apaga_emprego');
					}
					  } 
				
				function JsValorDelete(cod,cod2)
				   {
					if (confirm ("Tem certeza que deseja excluir esse valor?"))
				  {
				  document.location.href('../acoes/funcionarios_acao.asp?cod='+cod+'&cod_valor='+cod2+'&acao=apaga_valor');
					}
					  } 
				
					  
				//function JsContatoDelete(cod){
					//if (confirm ("Tem certeza que deseja excluir o contato?")){
				  //document.location.href('empresa/acoes/funcionarios_acao.asp?cod='+cod+'&acao=excluir');
					//}
					  //}
			</SCRIPT>