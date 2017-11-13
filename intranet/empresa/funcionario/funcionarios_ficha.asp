  <!--#include file="../../includes/util.asp"-->
  <!--#include file="../../includes/javascripts.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--#include file="../../includes/inc_area_restrita.asp"-->

<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<%
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
		strnm_tit_eleitor = rs("nm_tit_eleitor")
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
<style type="text/css" media="print">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.inputs { background-color: #cdcdcd; font: 10px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; }
.no_print{ 
	visibility:hidden; 
	display: none;}
.ok_print{
	visibility:visible;
	display:block;}
#frame{border:0px inset;
	position:relative;
	margin: 0px;
	padding: 6px;
	text-align: left;}
table{background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 200px;
	font-family: verdana;
	font-size:11px;
	text-decoration:none;}
-->
</style>
<style type="text/css" media="screen">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.cabecalho {color: #000000;font-family: verdana;font-size:14px;}
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.inputs { background-color: #cdcdcd; font: 10px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #cdcdcd; }
.ok_print{ 
	visibility:hidden; 
	display: none;}
.no_print{
	visibility:visible;
	display:block;}
#frame{border:0px inset;
	position:relative;
	margin: 0px;
	padding: 6px;
	text-align: left;}
table{background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 200px;
	font-family: verdana;
	font-size:11px;
	text-decoration:none;}
-->
</style>

<body>
	<table border="1" align="center" style="border:0px solid black; border-collapse:collapse;">
		<tr>
			<td colspan="4" align="center" class="cabecalho" style="background-color:gray; color: white;"><b>FICHA CADASTRAL DO COLABORADOR </b></td>
		</tr>
			
		<tr>
			<td rowspan="4" align="center" valign="middle"><%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%else%><img src="../../foto/funcionarios/<%=strfoto%>" alt="" name="<%=strfoto%>" id="<%=strfoto%>" width="73" border="0"><%end if%></td>
			<td colspan="2" style="background-color:silver; color: white;" align="center"><b>DADOS PESSOAIS</b></td>
			<td style="background-color:silver; color: white;" align="center"><b>MATRICULA</b></td>
		</tr>
		<tr>
			<td colspan="2">
				<b style="font-size:12px;"><%=strnome %></b><br>
				<%=strnm_endereco%>, <%=strnr_numero%><br>
				<%=strnm_complemento%> <%if strnm_complemento  <> "" Then%> - <%end if%><%=strbairro%> <br>
				<%=strnm_cidade%> - <%=strnm_estado%><br>
				<%=strnm_cep%></td>
			<td style="font-size:20px;" align="center"><b><%=strmatricula%></b></td>
		</tr>
		<tr><td colspan="3" align="center">&nbsp;<img src="../../imagens/ic_print_view.gif" alt="" width="24" height="26" border="0" onclick="visualizarImpressao();" class="no_print"></td></tr><tr>
		<tr>
			<td><b>Nasc.:</b> <%=strdt_nascimento%></td>
			<td ><b>Sexo</b>: <%if str_sexo = 1 then%>Masculino<%elseif str_sexo = 2 then%>Feminino<%end if%></td>
			<td><b>Estado Civil</b>: 
			<%strsql = "Select * From TBL_estado_civil Where cd_codigo="&str_estado_civil&""
				Set rs_est_civil = dbconn.execute(strsql)
				if not rs_est_civil.EOF Then
					response.write(rs_est_civil("nm_estado_civil"))
				end if
					rs_est_civil.close
					Set rs_est_civil = nothing%></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td><b>RG.:</b> <%=strnm_rg%></td>
			<td><b>Expedidor:</b> <%=strnm_rg_expedidor%></td>
			<td><b>CPF:</b> <%=strnm_cpf%></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2"><b>Carteira Profissional:</b> <%=str_ctps%> <b>Série:</b> <%=str_ctps_serie%></td>
			<td><b>PIS:</b> <%=str_pis%></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2"><b>Tit. Eleitor:</b> <%=strnm_tit_eleitor%></td>
			<td colspan="2"><b>Zona:</b> <%=strnr_tit_eleitor_zona%> &nbsp; <b>Secção:</b> <%=strnr_tit_eleitor_seccao%></td>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
			<!--td rowspan="2" align="center"><b>FILIAÇÃO</b></td-->
			<td colspan="2"><b>Mãe</b>: <%=strmae%></td>
		<!--/tr><tr-->
			<td colspan="2"><b>Pai</b>: <%=strpai%></td>
		</tr>
		<tr>
			<td colspan="4"><b>Nome do Cônjuge</b>: <%=str_conjuge%></td>
		</tr>									
		<!--tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="5"></td></tr>
		<tr>
			<td colspan="4"><b>Endereço alternativo</b>: <%=str_endereco_alt%></td>
		</tr-->
		<td colspan="4">&nbsp;</td>
		<tr>
			<td colspan="4">
				<table border="1" style="border:1px solid gray; border-collapse:collapse;">
					<tr>
						<td colspan="4" style="background-color:silver; color: white;" align="center"><b>CONTATOS</b></td>			
					</tr>
					<%strsql = "Select * From TBL_contato Where cd_origem=4 and id_origem='"&str_funcionario&"' order by cd_categoria"
						Set rs_contato = dbconn.execute(strsql)
						
						'if not rs_contato.EOF then
						while not rs_contato.EOF
							cd_contato = rs_contato("cd_codigo")
							nm_contato_nome = rs_contato("nm_nome")
							cd_contato_categoria = rs_contato("cd_categoria")
							nm_cargo = rs_contato("nm_cargo")
							cd_contato_ddd = rs_contato("cd_ddd")
							nm_contato_numero = rs_contato("nm_numero")
							nm_contato_obs = rs_contato("nm_obs")	
								if int(cd_contato_categoria) = 5 then
									cabeca_contato = 0
								end if%>
					
							<%if cabeca_contato = 0 Then%>
							<tr bgcolor="gray" style="color:white;">													
								<td><img src="../../imagens/px.gif" width="79" height="1"></td>								
								<td>&nbsp;<b>Tel/Cel/Email</b><br><img src="../../imagens/px.gif" width="125" height="1"></td>
								<%if cd_contato_categoria = 5 then%>
								<td>&nbsp;Nome<br><img src="../../imagens/px.gif" width="200" height="1"></td>
								<%end if%>
								<td>&nbsp;<b>Observações</b><br><img src="../../imagens/px.gif" width="200" height="1"></td>
								<%if cd_contato_categoria <> 5 then%>
								<td><img src="../../imagens/px.gif" width="200" height="1"></td>
								<%end if%>
							</tr>
							<%end if%>
							
							<tr>
								<%strsql = "Select * From TBL_contato_cat Where cd_codigo='"&cd_contato_categoria&"'"
								Set rs_contato_cat = dbconn.execute(strsql)%>
								<%Do While Not rs_contato_cat.eof
									categoria_contato = rs_contato_cat("categoria_contato")
									cd_contato_cat = rs_contato_cat("cd_codigo")
								rs_contato_cat.movenext
								loop
								rs_contato_cat.close
								Set rs_contato_cat = nothing%>													
								<td <%if cd_contato_categoria = 5 then%> style="color:red; font-weight:bold;"<%end if%>>&nbsp;<%=categoria_contato%></td>
								<td>&nbsp;<%if cd_contato_ddd <> "" Then response.write(cd_contato_ddd &" - ")%> <%=nm_contato_numero%></td>
								<td>&nbsp;<%=nm_contato_obs%></td>
								<%if cd_contato_categoria = 5 then%>
								<td>&nbsp;<%=nm_contato_nome%> - <%=nm_cargo%></td>
								<%else%>
								<td>&nbsp;</td>
								<%end if%>
							</tr>
					<%cabeca_contato = cabeca_contato + 1
					
					
					rs_contato.movenext
					wend%>
				</table>
			</td>
		</tr>							
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
			<td colspan="4">
				<table border="1" style="border:1px solid gray; border-collapse:collapse;">
					<tr>
						<td colspan="6" style="background-color:silver; color: white;" align="center"><b>DEPENDENTES</b></td>			
					</tr>
					<%strsql = "Select * From TBL_dependente Where cd_funcionario='"&str_funcionario&"' order by cd_parentesco,dt_nascimento,nm_nome"
						Set rs_dependente = dbconn.execute(strsql)
						
					if rs_dependente.EOF Then%>
						<td colspan="6" style="width:300px; font-weight:bold;" align="center">Não Possui dependentes</td>
					<%end if
					while not rs_dependente.EOF
						cd_dependente = rs_dependente("cd_codigo")
						nm_dependente_nome = rs_dependente("nm_nome")
						cd_dependente_parentesco = rs_dependente("cd_parentesco")
						dt_dependente_nascimento = rs_dependente("dt_nascimento")
						nm_dependente_obs = rs_dependente("nm_obs")
					
					conta_dependentes = conta_dependentes + 1
					if cd_dependente_parentesco = 1 then
						conta_filhos = conta_filhos + 1
					end if%>
					<%if cabeca_dependente = "" Then%>
					<tr bgcolor="#808080" style="color:white;">
						<td><img src="../../imagens/px.gif" width="15" height="1"></td>
						<td><b>Nome</b><br><img src="../../imagens/px.gif" width="235" height="1"></td>
						<td><b>Parent.</b><br><img src="../../imagens/px.gif" width="60" height="1"></td>
						<td><b>Nasc.</b><br><img src="../../imagens/px.gif" width="75" height="1"></td>
						<td><b>Idade</b><br><img src="../../imagens/px.gif" width="60" height="1"></td>
						<td><b>Observações</b><br><img src="../../imagens/px.gif" width="154" height="1"></td>											
					</tr>
					<%end if%>
					<tr>
						<td><b><%=conta_dependentes%></b></td>
						<td><%=nm_dependente_nome%></td>											
						<%strsql = "Select * From TBL_parentesco Where cd_codigo="&cd_dependente_parentesco&""
						Set rs_dependente_parent = dbconn.execute(strsql)
						
						While Not rs_dependente_parent.eof
							nm_dependente_parentesco = rs_dependente_parent("nm_parentesco")
						rs_dependente_parent.movenext
						wend
						'rs_dependente_parent.close
						'Set rs_dependente_parent = nothing%>
						
						<td><%=nm_dependente_parentesco%></td>
						<td><%=zero(day(dt_dependente_nascimento))&"/"&zero(month(dt_dependente_nascimento))&"/"&year(dt_dependente_nascimento)%></td>
					<%idade = datediff("m",dt_dependente_nascimento,now)
						if idade >= 12 then
							if idade = 12 then
								tempo = "ano"
							else
								tempo = "anos"
							end if
							idade = int(idade/12)&" "&tempo
						else 
							if idade <= 1 then
								tempo = "mes"
							else
								tempo = "meses"
							end if
							idade = idade&" "&tempo
						end if%>
						<td><%=idade%></td>
						<td><%=nm_dependente_obs%></td>
					</tr>
					<%cabeca_dependente = cabeca_dependente + 1
					'conta_dependentes = conta_dependentes + 1
					rs_dependente.movenext
					wend
					if cd_dependente <> "" then%>
					<tr><td colspan="7" bgcolor="silver" style="color:white;"><b>Quantidade de filhos: <%=conta_filhos%></b></td></tr>
					<tr><td colspan="7" style="color:gray;">Quantidade Total de dependentes: <b><%=conta_dependentes%></b></td></tr>
					<%end if%>
				</table>
			</td>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
			<td colspan="4" style="background-color:silver; color: white;" align="center"><b>DADOS PROFISSIONAIS</b></td>			
		</tr>
		<tr>
			<td><b><%strsql = "Select * From TBL_cons_profissional Where cd_codigo = '"&str_conspro&"'"
 					Set rs_conspro = dbconn.execute(strsql)
					While not rs_conspro.EOF
						cod_conspro = rs_conspro("cd_codigo")
						nm_conspro = rs_conspro("nm_conspro")%>												
						<%=trim(nm_conspro)%>:
					<%rs_conspro.movenext
					wend%></b> <%=str_numreg%></td>
			<td><b>Tipo:</b> <%if  str_rgpro_concessao = "" then 
								elseif str_rgpro_concessao = 1 then
									response.write("Definitivo")
								elseif str_rgpro_concessao = 2 then
									response.write("Provisório")
								end if%></td>
			<td valign="top"><b>Situação:</b> <%if str_rgpro_status = 1 then
													response.write("Ativo")
												elseif str_rgpro_status = 2 then
													response.write("Inativo")
												end if%></td>
			<td valign="top" rowspan="2"><b>Obs.:</b> <%=str_rgpro%></td>
		</tr>
		<tr>
			<td valign="top"><b>Inscr.:</b> <%=str_dt_rgproinscr%></td>
			<td valign="top"><b>Pagto.:</b> <%=str_dt_rgpropag%></td>
			<td valign="top"><b>Val.:</b> <%=str_dt_rgproval%></td>
			<!--td><input type="text" name="dia_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="mes_valid" size="3" maxlength="2" class="inputs">/<input type="text" name="ano_valid" size="5" maxlength="4" class="inputs"></td-->
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
			<td colspan="4" style="background-color:silver; color: white;" align="center"><b>SAÚDE</b></td>			
		</tr><tr>
			<td colspan="4" style="background-color:gray; color:white;"><b>Carteira de vacinação</b></td>
		</tr>
		<tr>
			<td><b>Hepatite B</b></td>
			<td colspan="3">
				<%if int(strnr_hepatite_b) = 3 then
					ckhepatite_b3 = "X"
					ckhepatite_b2 = "X"
					ckhepatite_b1 = "X"
				elseif int(strnr_hepatite_b) = 2 then
					ckhepatite_b3 = "&nbsp;&nbsp;"
					ckhepatite_b2 = "X"
					ckhepatite_b1 = "X"
				elseif int(strnr_hepatite_b) = 1 then
					ckhepatite_b3 = "&nbsp;&nbsp;"
					ckhepatite_b2 = "&nbsp;&nbsp;"
					ckhepatite_b1 = "X"
				else
					ckhepatite_b3 = "&nbsp;&nbsp;"
					ckhepatite_b2 = "&nbsp;&nbsp;"
					ckhepatite_b1 = "&nbsp;&nbsp;"
				end if%>
				<b>(</b><%=ckhepatite_b1%><b>) 1ª dose</b> &nbsp; 
				<b>(</b><%=ckhepatite_b2%><b>) 2ª dose</b> &nbsp;
				<b>(</b><%=ckhepatite_b3%><b>) 3ª dose</b></td>
		</tr>
		<tr>
			<td><b>Dupla Adulto</b></td>
			<td colspan="3">
				<%if int(strnr_dupla_adulto) = 3 then
					ckdp_adulto_b3 = "X"
					ckdp_adulto_b2 = "X"
					ckdp_adulto_b1 = "X"
				elseif int(strnr_dupla_adulto) = 2 then
					ckdp_adulto_b3 = "&nbsp;&nbsp;"
					ckdp_adulto_b2 = "X"
					ckdp_adulto_b1 = "X"
				elseif int(strnr_dupla_adulto) = 1 then
					ckdp_adulto_b3 = "&nbsp;&nbsp;"
					ckdp_adulto_b2 = "&nbsp;&nbsp;"
					ckdp_adulto_b1 = "X"
				Else
					ckdp_adulto_b3 = "&nbsp;&nbsp;"
					ckdp_adulto_b2 = "&nbsp;&nbsp;"
					ckdp_adulto_b1 = "&nbsp;&nbsp;"
				end if%>
				<b>(</b><%=ckdp_adulto_b1%><b>) 1ª dose</b> &nbsp; 
				<b>(</b><%=ckdp_adulto_b2%><b>) 2ª dose</b> &nbsp;
				<b>(</b><%=ckdp_adulto_b3%><b>) 3ª dose</b> &nbsp; &nbsp; 
				<b>Val.:</b> <%=strdt_dupla_adulto_validade%></td>
		</tr>
		<tr>
			<td><b>SCR*</b></td>
			<td><b>(</b><%if int(strnr_scr) = 1 then%>X<%else%>&nbsp;&nbsp;<%end if%><b>) SIM </b>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; 
				<b>(</b><%if int(strnr_scr) = "" OR int(strnr_scr) = 0 then%>X<%else%>&nbsp;&nbsp;<%end if%><b>) Não </b></td>	
			<td colspan="2"><span style="font-size:9px; color:gray;">Sarampo, Caxumba, Rubeóla.</span></td>							
		</tr>
		<tr><td colspan="4"><br><img src="../imagens/blackdot.gif" width="100%" height="1"></td></tr><tr>
		<tr>
			<td colspan="2"><b>Exame médico realizado em: </b><%=strdt_exame_medico%></td>
			<td colspan="2"><b>Validade: </b><%if strdt_exame_medico <> "" Then%><%=day(strdt_exame_medico)&"/"&month(strdt_exame_medico)&"/"&year(strdt_exame_medico)+1%><%end if%></td>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr><tr>
		<tr>
			<td colspan="4">
				<table border="1" style="border:1px solid gray; border-collapse:collapse;">
					<tr><td colspan="5" align="center" style="background-color:silver; color: white;"><b>ESCOLARIDADE</b></td></tr>
					<tr bgcolor="#808080" style="color:white;">
						<td align="center"><b>Instituição</b><br><img src="../../imagens/px.gif" alt="" width="220" height="1" border="0"></td>
						<td align="center"><b>Grau</b><br><img src="../../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
						<td align="center"><b>Curso</b><br><img src="../../imagens/px.gif" alt="" width="170" height="1" border="0"></td>
						<td align="center"><b>Andamento</b><br><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
						<td align="center"><b>Término</b><br><img src="../../imagens/px.gif" alt="" width="65" height="1" border="0"></td>
					</tr>
					<%strsql = "Select * From TBL_escolaridade Where cd_funcionario='"&str_funcionario&"' order by cd_ensino_grau"
						Set rs_escolaridade = dbconn.execute(strsql)
						
					'if rs_escolaridade.EOF Then%>
						<!--td colspan="100%" style="width:300px; font-weight:bold;" align="center">&nbsp;</td-->
					<%'end if
					while not rs_escolaridade.EOF
						cd_escolaridade = rs_escolaridade("cd_codigo")
						cd_escolaridade_grau = rs_escolaridade("cd_ensino_grau")
						cd_escolaridade_andamento = rs_escolaridade("cd_ensino_andamento")
						nm_escolaridade_instituicao = rs_escolaridade("nm_instituicao")
						nm_escolaridade_curso = rs_escolaridade("nm_curso")
						nm_escolaridade_obs = rs_escolaridade("nm_obs")
						dt_escolaridade_termino = rs_escolaridade("dt_termino")
					
					'conta_dependentes = conta_dependentes + 1
					'if cd_dependente_parentesco = 1 then
					'	conta_filhos = conta_filhos + 1
					'end if%>					
					<tr>
						<td><%=nm_escolaridade_instituicao%></td>
						<td><%xsql = "Select * From TBL_escolaridade_grau where cd_codigo='"&cd_escolaridade_grau&"'"
										Set rs_grau = dbconn.execute(xsql)
										if not rs_grau.EOF then
										cd_grau = rs_grau("cd_codigo")
										nm_grau = rs_grau("nm_grau")%>
											<%=nm_grau%>
										<%rs_grau.movenext
										end if%></td>
						<td><%=nm_escolaridade_curso%></td>
						<td><%xsql = "Select * From TBL_escolaridade_andamento where cd_codigo='"&cd_escolaridade_andamento&"'"
										Set rs_andamento = dbconn.execute(xsql)
										if not rs_andamento.EOF then
										cd_andamento = rs_andamento("cd_codigo")
										nm_ensino_andamento = rs_andamento("nm_ensino_andamento")%>
											<%=nm_ensino_andamento%>
										<%rs_andamento.movenext
										end if%></td>
						<td><%=dt_escolaridade_termino%></td>										
					</tr>
					<%rs_escolaridade.movenext
					wend%>
					
				</table>
			</td>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr>
		<tr>
			<td colspan="4">
				<table border="1" style="border:1px solid gray; border-collapse:collapse;">
					<%strsql = "Select * From TBL_emprego_anterior Where cd_funcionario='"&str_funcionario&"' order by dt_entrada"
						Set rs_emprego = dbconn.execute(strsql)
						
						while not rs_emprego.EOF
						
						cd_emprego = rs_emprego("cd_codigo")
						nm_emprego_empresa = rs_emprego("nm_empresa")
						nm_emprego_cargo = rs_emprego("nm_cargo")
						dt_emprego_entrada = rs_emprego("dt_entrada")
						dt_emprego_saida = rs_emprego("dt_saida")
						nm_emprego_obs = rs_emprego("nm_obs")
						
						conta_linha = conta_linha + 1%>
					<%if cabeca_emprego = "" Then%>
					<tr><td colspan="6" align="center" style="background-color:silver; color: white;"><b>EXPERIÊNCIA</b></td></tr>
					<tr bgcolor="#808080" style="color:white;">
						<td><img src="../../imagens/px.gif" alt="" width="20" height="1" border="0"></td>
						<td align="center"><b>Emprego Anterior</b><br><img src="../../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
						<td align="center"><b>Função</b><br><img src="../../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
						<td align="center"><b>Admissão</b><br><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
						<td align="center"><b>Recisão</b><br><img src="../../imagens/px.gif" alt="" width="75" height="1" border="0"></td>
						<td align="center"><b>Observação</b><br><img src="../../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
					</tr>
					<%end if%>
					<tr>
						<td>&nbsp;<%=conta_linha%></td>
						<td>&nbsp;<%=nm_emprego_empresa%></td>											
						<td>&nbsp;<%=nm_emprego_cargo%></td>
						<td>&nbsp;<%=dt_emprego_entrada%></td>
						<td>&nbsp;<%=dt_emprego_saida%></td>
						<td>&nbsp;<%=nm_emprego_obs%></td>
					</tr>
					<%cabeca_emprego = cabeca_emprego + 1
					rs_emprego.movenext
					wend%>
				</table>
			</td>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr><tr>
		
		<tr><td colspan="4" align="center" style="background-color:silver; color: white;"><b>BENEFICIOS</b></td></tr>
		<tr>
			<td><b>SPTRANS:</b> <%=str_sptrans%></td>
			<td><b>BOM:</b> <%=str_cmt_bom%></td>
			<td><b>VR:</b> <%=str_vr%></td>
			<td><b>Cesta:</b> <%=str_vale_alimentacao%></td>
		</tr>
		
		<tr>
			<td><b>Assist. Médica:</b> <%=str_assistencia_medica_matricula%></td>
			<td><b>Co-part.:</b> <%=str_assistencia_medica_coparticipacao%>%</td>
			<td><b>Assist. Odonto:</b> <%=str_assistencia_odonto_matricula%></td>
			<td><b>Co-part.:</b> <%=str_assistencia_odonto_coparticipacao%></td>
		</tr>		
		<tr>
			<td><b>Banco:</b> <%=str_banco%></td>
			<td><b>Agência:</b> <%=str_banco_ag%></td>
			<td><b>Conta</b> <%=str_banco_conta%></td>
			<td>&nbsp;</td>
		</tr>
		<tr><td colspan="4">&nbsp;</td></tr>				
		<tr>
			<td colspan="4" align="center" style="background-color:silver; color: white;"><b>CONTRATO</b></td>			
		</tr>
		<%strsql = "Select * From VIEW_funcionario_contrato_lista Where cd_funcionario = "&str_funcionario&" order by dt_admissao desc"
				Set rs_contrato = dbconn.execute(strsql)%>
		<%while not rs_contrato.EOF
			cd_funcionario = rs_contrato("cd_funcionario")
			'cd_regime_trabalho = rs_contrato("cd_codigo")
			nm_regime_trab = rs_contrato("nm_regime_trab")
			dt_admissao = rs_contrato("dt_admissao")
			dt_demissao = rs_contrato("dt_demissao")%>
		<tr>
			<td colspan="2"><b>Contrato:</b> <%=nm_regime_trab%></td>
			<td><b>Admissão:</b> <%=zero(day(dt_admissao))%>/<%=zero(month(dt_admissao))%>/<%=year(dt_admissao)%></td>
			<td><b>Recisão:</b> <%if dt_demissao <> "" Then%><%=zero(day(dt_demissao))%>/<%=zero(month(dt_demissao))%>/<%=year(dt_demissao)%><%end if%></td>			
		</tr>
			<%nr_contrato = nr_contrato + 1
			rs_contrato.movenext
			wend%>
		<tr>
			<td colspan="2"><b>Cargo:</b>
			<%strsql = "SELECT * FROM VIEW_funcionario_cargo where cd_funcionario = '"&str_funcionario&"' order by dt_atualizacao desc"
			Set rs_func = dbconn.execute(strsql)
				if Not rs_func.eof then%>
					<%=rs_func("nm_cargo")%>
				<%end if
				
				rs_func.close
				Set rs_func = nothing%></td>								
			<td><b>Unidade:</b>
			<%strsql ="SELECT * From View_funcionario_unidade where cd_funcionario = "&str_funcionario&" and cd_suspenso <> 1 order by dt_atualizacao desc"
			Set rs_unidade = dbconn.execute(strsql)
				
				if not rs_unidade.EOF then%>
					<%=rs_unidade("nm_unidade")%>
				<%end if%>				 
				<%rs_unidade.close
				Set rs_unidade = nothing%>
			</td>
			<td>&nbsp;</td>
		</tr>
		<%xsql = "Select * From TBL_funcionario_horario Where cd_funcionario = '"&strcod&"' and cd_suspenso <> 1 ORDER BY dt_atualizacao DESC"
		Set rs_hora = dbconn.execute(xsql)
			if not rs_hora.EOF then
				hr_entrada = rs_hora("hr_entrada")
				hr_saida = rs_hora("hr_saida")
				nm_intervalo = rs_hora("nm_intervalo")
			end if
		  %>														
		<tr>
			<td><b>Entrada:</b> <%=zero(hour(hr_entrada))&":"&zero(minute(hr_entrada))%><br><img src="../../imagens/px.gif" width="140" height="1"></td>
			<td><b>Intervalo:</b> <%=nm_intervalo%><br><img src="../../imagens/px.gif" width="140" height="1"></td>
			<td><b>Saída:</b> <%=zero(hour(hr_saida))&":"&zero(minute(hr_saida))%><br><img src="../../imagens/px.gif" width="140" height="1"></td>				
			<td><img src="../../imagens/px.gif" width="140" height="1"></td>
		</tr>
		<tr><td colspan="4"><img src="../../imagens/px.gif" width="620" height="10"></td></tr>
	</table>
						
						<br><hr align="center" width="620">
						<br>
							
					
									 								
							<%'End if%>
</body>	