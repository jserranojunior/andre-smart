<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->

<%
session_cd_usuario = session("cd_codigo")

strcod = request("cod")
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strbusca = request("busca")
stracao = request("acao")

dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")

if dt_dia = "" Then dt_dia = day(now) end if
if dt_mes = "" Then dt_mes = month(now) end if
if dt_ano = "" Then dt_ano = year(now) end if

dia_final = ultimodiames(dt_mes,dt_ano)


			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				'ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				ordem_funcionarios = "cd_contrato,nm_nome"
			end if

	'xsql = "up_funcionario_rh_lista @dt_data='"&dt_ano&dt_mes&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
	'	Set rs_func = dbconn.execute(xsql)
	'	response.write(func_ativos)
%>


<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<!--#include file="../../js/ajax.js"-->
<script language="Javascript">
<!-- 
//-------------------------------------------------------------------

function adiciona_valores(a,b,c,d,e,valores_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	e=e;
	valores_total=valores_total;
			
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
			while (e.indexOf(' ') != -1) {
	 		e = e.replace(' ','%20');}
			
			while (valores_total.indexOf('$$') != -1) {
	 		valores_total = valores_total.replace('$$','$');}
						
		if (valores_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	valores_total = valores_total + separador
		if (a != ""){
			valores_total = valores_total + a + '#' + c + '#' + d + '#' + e + '#';
			document.form.valores_result.value=a;} //adiciona				
						
		if (b != ""){
			valores_total = valores_total.replace(b,'');
			document.form.valores_total.value=valores_total;} //subtrai			

	document.form.valores_total.value=(valores_total.replace("$$", "$"));	
	document.form.valores_total.value=(valores_total); //c
	document.form.cd_valor_tipo.value="";
	document.form.nr_valor.value="";
	document.form.dt_valor_atualizacao.value="";
	document.form.nm_valor_obs.value="";
	
	document.form.cd_valor_tipo.focus();
	//chama Ajax	
	{
		var oHTTPRequest_valores = createXMLHTTP(); 
		oHTTPRequest_valores.open("post", "valores/ajax/ajax_valores_lista.asp", true);
		oHTTPRequest_valores.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		oHTTPRequest_valores.onreadystatechange=function(){
		if (oHTTPRequest_valores.readyState==4){
	    document.all.divValores_lista.innerHTML = oHTTPRequest_valores.responseText;}}
	    oHTTPRequest_valores.send("valores_total=" + form.valores_total.value);
	}
 
}



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

					<% If strmsg <> "" then%>
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
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print" style="">
						<form name="form" action="../acoes/funcionarios_acao.asp"  method="POST"  onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="cd_sessao" value="<%=session_cd_usuario%>">
						<input type="hidden" name="jan" value="1">
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
								<td colspan="7" align="center" class="cabecalho" style="background-color:black; color:white;"><b>ADM - FOLHA DE PAGAMENTO <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Mês de competência:</b></td>
								<td align="center" style="background-color:silver;" colspan="2">
									<select name="dt_mes">
										<%for i = 1 to 12%>
											<option value="<%=i%>" <%if i = dt_mes then%>SELECTED<%end if%>><%=mesdoano(dt_mes)%></option>
										<%next%>
									</select> / <input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4"></td>									
							</tr>
							<tr>
								<td colspan="1"><img src="../../imagens/blackdot.gif" width="150" height="1">&nbsp;</td>
								<td colspan="2"><img src="../../imagens/blackdot.gif" width="250" height="1">&nbsp;</td>
							</tr>
							<%	xsql = "up_funcionario_rh_lista_individual @cd_funcionario="&strcod&", @dt_data='"&dt_ano&dt_mes&"', @dt_atualizacao = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @ordem_funcionarios='"&ordem_funcionarios&"'"
								Set rs_func = dbconn.execute(xsql)
									
									if not rs_func.EOF then
										cd_funcionario = rs_func("cd_funcionario")
										nm_funcionario = rs_func("nm_nome")
										cd_matricula = rs_func("cd_matricula")
										nm_foto = rs_func("nm_foto")
										
										'expediente = rs_func("expediente")
										
										'nr_salminimo = rs_func("nr_salminimo")
										nr_salario = rs_func("nr_salario")
										'dt_admissao = zero(day(rs_func("dt_admissao")))&"/"&zero(month(rs_func("dt_admissao")))&"/"&year(rs_func("dt_admissao"))
										'dt_demissao = zero(day(rs_func("dt_demissao")))&"/"&zero(month(rs_func("dt_demissao")))&"/"&year(rs_func("dt_demissao"))
										
										admissao = rs_func("admissao")
										demissao = rs_func("demissao")
										
										cd_contrato = rs_func("cd_contrato")
										
										'cd_unidade = rs_func("cd_unidade")
										'nm_unidade = rs_func("nm_unidade")
										'nm_sigla = rs_func("nm_sigla")
										'cd_hospital = rs_func("cd_hospital")
										
									'	idade_depend = rs_func("idade_depend")
										
									end if%>
										<tr class="no_print">
								<td rowspan="2" align="center" valign="top" style="color:#7d7d7d;">									
										<img src="../../foto/funcionarios/<%=nm_foto%>" alt="" name="<%=nm_foto%>" id="<%=strfoto%>" width="73" border="0"><br>
									<%if strcod <> "" then%>
										<a href="#" onClick="window.open('empresa/janelas/funcionarios_foto.asp?cd_codigo=<%=strcod%>','upload','width=500,height=550')" alt="Inserir/mudar foto">(Alterar foto)</a>
									<%end if%>
								</td>
								<%'if 
								'end if%>
								<td class="no_print"><b>Matrícula+</b></td>
								<td align="center" class="no_print"><b>Expediente</b></td>
							</tr>
							<tr class="no_print">
								<td><b style="font-size:20px;"><%=cd_matricula%></b></td>
								<td align="center"><%cd_expediente = expediente/60
													if cd_expediente <= 6 then
														cd_expediente = "6 horas"
													elseif cd_expediente <= 11 then
														cd_expediente = "8 horas"
													elseif cd_expediente >= 12 then
														cd_expediente = "12/36"
													end if%>
													<b><%=cd_expediente%></b></td>
							</tr>
							<tr>
								<td  align="right" colspan=""><b style="color:red;">Nome</b></td>
								<td colspan="2"><b><%=nm_funcionario %>&nbsp;</b></td>
							</tr>
							<tr>
								<td align="center" colspan="3"><b>Valores Fixos</b> <%=cd_unidade%></td>								
							</tr>
							<tr>
								<td align="right"><b>Sal&aacute;rio</b></td>
								<td colspan="2"><%=nr_salario%></td>
							</tr>
							<tr>
								<td align="right"><b>Insalubridade</b></td>
								<td colspan="2"><%if cd_hospital = 1 then
													'xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=7 AND dt_atualizador<='"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
													'	Set rs_valor = dbconn.execute(xsql)
													'	if not rs_valor.EOF then
													'	'	dt_insalubr_indv = rs_valor("dt_atualizacao")
													'	'	insalubr_indv = int(year(dt_insalubr_indv)&zero(month(dt_insalubr_indv))&zero(day(dt_insalubr_indv)))
													'		insalubridade = rs_valor("nr_valor")
													'		response.write(insalubridade)
													'	end if
													insalubridade = nr_salminimo * 0.1
													response.write(formatnumber(insalubridade,2))
												else
													response.write("0")
												end if%></td>
							</tr>
							
							<tr>
								<td align="right"><b>Adicional Noturno</b></td>
								<%xsql = "Select * From VIEW_funcionario_valores Where cd_funcionario='"&strcod&"' AND cd_tipo=8 AND dt_atualizacao between '"&dt_mes&"/1/"&dt_ano&"' AND '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
									Set rs_valor = dbconn.execute(xsql)
									if not rs_valor.EOF then
										'dt_ad_noturno_indv = rs_valor("dt_atualizacao")
										'ad_noturno_indv = int(year(dt_ad_noturno_indv)&zero(month(dt_ad_noturno_indv))&zero(day(dt_ad_noturno_indv)))
										qtd_dia_noturno = rs_valor("nr_valor")
									end if%>
								<td><%ad_noturno = (((nr_salario*0.40)/180)*7)*qtd_dia_noturno
											if not ad_noturno = 0 then
											'	response.write(formatnumber(ad_noturno,2))
												'ad_noturno_dsr = formatnumber((ad_noturno/25)*6.25,2)
											'else
											'	ad_noturno_dsr = 0
											end if
											'response.write(formatnumber(ad_noturno,2))%>&nbsp;</td>
								<td><%=int(qtd_dia_noturno)%> dia<%if qtd_dia_noturno > 1 then response.write("s")%>&nbsp;</td>
							</tr>
							<tr>
								<td align="right"><b>Aux&iacute;lio Creche</b></td>
								<td><%if idade_depend > 0 Then%>
										<%=formatnumber(idade_depend*valor_auxilio_creche_padrao,2)%>
									<%else%>
										0
									<%end if%></td>
							</tr>
							<tr>
								<td align="right"><b>Hora Extra</b></td>
							</tr>
							<tr>
								<td align="right"><b>Contribuição Sindical</b></td>
							</tr>
							<tr>
								<td align="right"><b>Cesta Básica</b></td>
								<td colspan="2"><%xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=10 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
											Set rs_refeic = dbconn.execute(xsql)
											if not rs_refeic.EOF then
												cesta_b = rs_refeic("nr_valor")
												response.write(formatnumber(cesta_b,2))
											end if%></td>
							</tr>
							<tr>
								<td align="right"><b>Vale Alimentação</b></td>
								<td colspan="2"><%xsql = "Select * From VIEW_funcionario_horario Where cd_funcionario='"&strcod&"' AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
										Set rs_hora = dbconn.execute(xsql)
											if not rs_hora.EOF then
												expediente = rs_hora("expediente")
													if abs(expediente) > 360 then
														xsql = "Select * From VIEW_funcionario_valores_padroes Where cd_tipo=2 AND dt_atualizacao <= '"&dt_mes&"/1/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
															Set rs_refeic = dbconn.execute(xsql)
															if not rs_refeic.EOF then
																valor_refeicao = rs_refeic("nr_valor")
															end if
													end if
												response.write(formatnumber(valor_refeicao))
											end if%></td>
							</tr>
							<tr>
								<td align="right"><b>Vale Transporte</b></td>
								<td colspan="2"><%if nr_salario <> "" Then
											v_trasnsp = nr_salario*0.06
											response.write(formatnumber(v_trasnsp))
										else
											response.write("0")
										end if%></td>
							</tr>
							<tr>
								<td align="right"><b>Base INSS</b></td>
								<td colspan="2"><%base_inss = abs(valor_salario) + abs(insalubridade) + abs(ad_noturno) +abs(horas_extras) + abs(horas_extras_dsr) + abs(ad_noturno_dsr)
									if not base_inss = 0 then response.write(formatnumber(base_inss,2))%></td>
							</tr>
							<tr>
								<td align="right"><b>INSS</b></td>
								<td colspan="2"><%xsql = "Select * From TBL_inss Where dt_atualizacao <= '"&dt_mes&"/"&ultimodiames(dt_mes,dt_ano)&"/"&dt_ano&"' ORDER BY dt_atualizacao DESC"
												Set rs_inss = dbconn.execute(xsql)
												if not rs_inss.EOF then
													faixa_1 = rs_inss("pri_faixa")
													faixa_2 = rs_inss("seg_faixa")
													faixa_3 = rs_inss("ter_faixa")
												end if
												
												if base_inss <= abs(faixa_1) then
													inss = base_inss * 0.08
													response.write(formatnumber(inss,2))
												elseif base_inss <= abs(faixa_2) then
													inss = base_inss * 0.09
													response.write(formatnumber(inss,2))
												elseif base_inss <= abs(faixa_3) then
													inss = base_inss * 0.11
													response.write(formatnumber(inss,2))
												end if%></td>
							</tr>
							
							<tr>
								<td>-------</td>
							</tr>
							
							<tr></tr>
							
							<tr class="no_print">
								<td colspan="3" valign="middle"><!--<A href="empresa.asp?tipo=lista"><img src="../../imagens/bt_lista.gif" alt="Listar" border="0" width="119" height="22" border="1"></a>-->
								<input type="image" src="../../imagens/bt.gif" alt="Confirmar" width="119" height="22" border="0">
								<%=stracao%></td>
							</tr>
							
							</form>
							<%If strcod <> "" then%>							
								<tr>
									<td colspan="3">&nbsp;
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
					  
				
			
			</SCRIPT>