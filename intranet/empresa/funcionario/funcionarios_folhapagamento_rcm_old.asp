<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../css/geral.htm"-->
<%
cd_user = session("cd_codigo")
jan = request("jan")
if jan = "" Then jan = "0"

str_ok = request("ok")

strcod = request("cod")
strcod = ""
strmsg = request("msg")
strpag = request("pag")
strtipo = request("tipo")
strcat = "enfermagem"'request("cat")


strbusca = request("busca")
stracao = request("acao")

strcd_unidade = request("cd_unidade")
	if strcd_unidade = "" Then	strcd_unidade = 0
	
mostra_faltas = request("mostra_faltas")
	if mostra_faltas = "" Then mostra_faltas = 0
	
cd_ordem = request("cd_ordem")
	if cd_ordem = "1" then
		nm_ordem = "nm_unidade,nm_nome"
	Elseif cd_ordem = "2" then
		nm_ordem = "nm_nome,nm_unidade"
	Elseif cd_ordem = "3" then
		'nm_ordem = "nr_faltas_justif,nr_faltas_injust"
		nm_ordem = "nr_faltas_justif,nr_faltas_injust"
		
	Elseif cd_ordem = "" AND strcat = "enfermagem"Then 
		nm_ordem = "nm_unidade,nm_nome"
		cd_ordem = "1"
	elseif cd_ordem = "" AND strcat = "empresa" Then 
		nm_ordem = "nm_nome,nm_unidade"
		cd_ordem = "2"
	
	end if

'***************************************************************
dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

dt_dia = request("dt_dia")
dt_mes = request("dt_mes")
dt_ano = request("dt_ano")
	if dt_dia = "" Then dt_dia = day(now) end if
	if dt_mes = "" Then dt_mes = month(now) 'end if
	if dt_ano = "" Then dt_ano = year(now) 'end if
	
	if dt_mes = "" AND dt_ano = "" Then
		dt_dia = day(now)
		dt_mes = month(now)
		dt_ano = year(now)
			if dt_dia > 15 then
				dt_dia = day(now)
				dt_mes = month(now)'+1
				dt_ano = year(now)
					if dt_mes = 13 then
						dt_mes = 1
						dt_ano = year(now)'+1
					end if
						if dt_dia > ultimodiames(dt_mes,dt_ano) then
							dt_dia = ultimodiames(dt_mes,dt_ano)
						end if
			end if
	end if
	
	mes_ant = dt_mes - 1
	ano_ant = dt_ano
		if dt_mes < 1 then
			dt_mes = 12
			ano_ant = dt_ano - 1
		end if 
	
dia_final = ultimodiames(dt_mes,dt_ano)

	mesanoatual=year(now)&zero(month(now))
	mesanosel=dt_ano&zero(dt_mes)
		
	mes_competencia = dt_mes' - 1
		'if mes_competencia < 1 then
		'	mes_competencia = 12
		'	ano_competencia = dt_ano - 1
		'else
			ano_competencia = dt_ano
		'end if
		final_competencia = ultimodiames(mes_competencia,ano_competencia)
'********************************************************************************

'*** Quantidade de dias da competencia ***
	'qtd_dias_comp = datediff("d","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia)'+1
	'qtd_dias_comp = datediff("d","1/9/2011","30/9/2011")

'*** Quantidade de domingos da competencia ***
	'qtd_semanas_comp = datediff("ww","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia,1)
	'qtd_domingos_comp = datediff("ww","1/"&mes_competencia&"/"&ano_competencia,final_competencia&"/"&mes_competencia&"/"&ano_competencia,1)
	
	
	for i = 1 to dia_final
		var_dia_semana = var_dia_semana&","&i
	next
	
'response.write("<br>"&dt_mes&"1/"&dt_ano&"<br>"&dia_final&"/"&dt_mes&"/"&dt_ano)

			if ordem_total <> "" Then
				ordem_funcionarios = ordem_total
			else
				'ordem_funcionarios = "cd_contrato,nm_nome,nm_sigla"
				ordem_funcionarios = "cd_contrato,nm_nome"
			end if
			
data_atual = dia_final&"/"&dt_mes&"/"&dt_ano%>
<!--include file="../../includes/feriados.asp"-->

<%qtd_dias_dsr = qtd_domingos_comp + qtd_feriados%>
<script language="javascript">
	function manipulacao_valor(cd_diario,dt_vencimento_anterior,dt_mes,dt_ano,acao,user){
		shipinfo = document.getElementById('seleciona_periodo');
		hoje = new Date
		mes_hoje = hoje.getMonth()+1;
			if (mes_hoje < 10 ) {
				mes_hoje = '0'+ mes_hoje;
				ano_hoje = hoje.getFullYear();}				
			data_hoje = ano_hoje +''+ (mes_hoje);
		 	data_informada = dt_ano +''+ dt_mes;
			/*if (data_informada<data_hoje){
				window.alert ("O mês selecionado já está fechado.");  return (false);}
			else{	*/
				window.open('financeiro/diario_cad3.asp?cd_diario='+cd_diario+'&dia_sel='+dt_vencimento_anterior+'&mes_sel='+dt_mes+'&ano_sel='+dt_ano+'&acao='+acao+'','Cadastro','width=500,height=280,scrollbars=1');
				/*}*/
	}
	
	function adiciona_pgto(a,lista_pgto,check1,check2){
		 lista_pgto = lista_pgto+','+a
		document.data.total_pgto.value=(lista_pgto);
		var el1=document.getElementById(check1);
			el1.style.display=(el1.style.display!='none'?'none':'');
		var el2=document.getElementById(check2);
			el2.style.display=(el2.style.display!='none'?'none':'');		
		}
	
	function subtrai_pgto(b,lista_pgto,check2,check1){
		lista_pgto = lista_pgto.replace(','+b,'');
		document.data.total_pgto.value=(lista_pgto);
		var el2=document.getElementById(check2);
			el2.style.display=(el2.style.display!='none'?'none':'');
		var el1=document.getElementById(check1);
			el1.style.display=(el1.style.display!='none'?'none':'');
		}
	
	function monta_pgto(lista_pgto){
		mywindow = window.open("financeiro/diario_pagamentos3.asp?lista_pagto="+lista_pgto+"&mes_sel=<%=dt_mes%>&ano_sel=<%=dt_ano%>", "Emissao_Cheques", "location=0,status=1,width=700,height=400,scrollbars=1, resizable=1");  
		mywindow.moveTo(0, 0); 
}	

	
</script>


					<table border="1" cellspacing="0" cellpadding="2" align="center" width="100" class="no_print">
						<%if session("cd_codigo") = 46 then%><tr><td colspan="100" align="center"><span style="font-size=8px;color:red;">funcionarios_folhapagamento_rcm.asp</span></td></tr><%end if%>
						<%if strfoto = "" then%><%strfoto = "Vdlap.gif"%><%end if%>
						<form name="data" action="<%if jan = 1 then%><%="funcionarios_folhapagamento_var"%><%else%><%=strcat%><%end if%>.asp"  method="GET">
							<input type="hidden" name="total_pgto" value="">
							<input type="hidden" name="tipo" value="<%=strtipo%>">
							<input type="hidden" name="cat" value="<%=strcat%>">
							<input type="hidden" name="cod" value="<%=strcod%>">
							<input type="hidden" name="jan" value="<%=jan%>">
							<input type="hidden" name="dia_sel" value="<%=dia_sel%>">
							<%if strcat = "enfermagem" then
								tit_doc = "ENFERMAGEM"
							Else
								tit_doc = "FOLHA DE PAGAMENTO"
							end if%>
							<tr>
								<td colspan="4" align="center" class="cabecalho" style="background-color:black; color:white;"><%'=dt_mes&"/"&dia_final&"/"&dt_ano%> <b>RCM - <%=tit_doc%> <%if strregime_trabalho <> "" Then%><%response.write("- "&strregime_trabalho&"."&strmatricula)%><%end if%></b></td>
								<!--td colspan="2" rowspan="4" align="center"><img src="foto/funcionarios/Vdlap.gif" alt="" name="Vdlap.gif" id="Vdlap.gif" width="73" border="0"></td-->
							</tr>
							<tr>
								<td align="center" style="background-color:silver;"><b>Período:</b><br><img src="../../imagens/px.gif" alt="" width="150" height="0" border="0"></td>
								<td align="left" style="background-color:silver;" colspan="3">
									<select name="dt_mes">
										<%for i = 1 to 12
											i_ant = i - 1
												if i_ant < 1 then i_ant = 12%>
											<option value="<%=i%>" <%if i = int(dt_mes) then%>SELECTED<%end if%>>21/<%=mesdoano(i_ant)%> - 20/<%=mesdoano(i)%></option>
										<%next%>
									</select> / <input type="text" name="dt_ano" value="<%=dt_ano%>" size="4" maxlength="4"></td>
							</tr>
							<!--tr>
								<td align="center" style="background-color:silver;">Mês de competência:</td>
								<td align="left" style="background-color:silver; color:red;"><b><%=ucase(mesdoano(mes_competencia))%>/<%=ano_competencia%></b></td>
							</tr-->
							<tr>
								<td align="center" style="background-color:silver;">Unidade</td>
								<td><%strsql ="TBL_unidades where cd_hospital > 0 and cd_status = 1"
								  	Set rs_uni = dbconn.execute(strsql)%>
									<select name="cd_unidade" class="inputs"> 
										<option value="0">Todas</option>
										<%Do While Not rs_uni.eof%>
										<option value="<%=rs_uni("cd_codigo")%>" <%if int(strcd_unidade) = rs_uni("cd_codigo") then response.write("SELECTED")%>><%=rs_uni("nm_Unidade")%></option>
										<%rs_uni.movenext
										loop
										rs_uni.close
										Set rs_uni = nothing%>
									</select></td>
							</tr>
							<!--tr>
								<td align="center" style="background-color:silver;">Faltas</td>
								<td style="background-color:silver;"><select name="mostra_faltas" class="inputs">
									<option value="0" <%if mostra_faltas = "" OR mostra_faltas = 0 then response.write("SELECTED")%> />Todos<br />
									<option value="1" <%if mostra_faltas = 1 then response.write("SELECTED")%> />Sem faltas<br />
									<option value="2" <%if mostra_faltas = 2 then response.write("SELECTED")%> />Com faltas
								</select></td>
							</tr-->
							<!--tr>
								<td align="center" style="background-color:silver;">Ordem</td>
								<td style="background-color:silver;"><select name="cd_ordem" class="inputs">
										<option value="1" <%if int(cd_ordem) = 1 then response.write("SELECTED")%>>Unidade</option>
										<option value="2" <%if int(cd_ordem) = 2 then response.write("SELECTED")%>>Alfabética</option>
										<option value="3" <%if int(cd_ordem) = 3 then response.write("SELECTED")%>>Faltas</option>
									</select></td>
							</tr-->
							<tr>
								<td align="left" colspan="1" style="background-color:silver;">&nbsp;</td>
								<td align="left" colspan="1" style="background-color:silver;"><input type="submit" name="ok" value="Visualizar"> &nbsp; &nbsp; <!--a href="javascript:void(0);" return false;" onClick="window.open('<%if jan = 1 then%><%else%>empresa/funcionario/<%end if%>funcionarios_folhapagamento_var_contabilidade.asp?dt_dia=<%=dt_dia%>&dt_mes=<%=dt_mes%>&dt_ano=<%=dt_ano%>','contab<%=dt_ano&zero(dt_mes)%>','location=0,status=1,width=750,height=600,scrollbars=1')">Contabilidade</a--></td></tr>
						</form>
					</table>
					
	<%'if str_ok = "Visualizar" then%>
					<br class="no_print">
					<table border="1" cellspacing="0" cellpadding="2" align="center" width="500" style="border-collapse:collapse;">					
			<!-- ---------------------------------------------------------------- -->				
							<tr bgcolor="#000000">
								<td colspan="17" align="center" class="cabecalho" style="background-color:black; color:white;"><b>RCM - <%=mesdoano(mes_competencia)%>/<%=ano_competencia%></b></td>
							</tr>
							
					<%num_cor = 1
					num_funcionario = 1
					outras_cond = " AND cd_contrato = 1 AND cd_unidade <> 27 "
					'subtotal_he_limit = 0
					'xsql = "up_funcionario_folhapagamento_geral @dt_comp_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @outras_cond='"&outras_cond&"', @nm_ordem='"&nm_ordem&"'"
					xsql = "up_funcionario_folhapagamento_geral @dt_comp_i='"&mes_ant&"/21/"&ano_ant&"', @dt_comp_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"', @dt_i = '"&dt_mes&"/1/"&dt_ano&"', @dt_f = '"&dt_mes&"/"&dia_final&"/"&dt_ano&"', @outras_cond='"&outras_cond&"', @nm_ordem='"&nm_ordem&"'"
					'xsql = "up_funcionario_rcm_geral @dt_comp_i='"&mes_ant&"/21/"&ano_ant&"', @dt_comp_f='"&mes_competencia&"/20/"&ano_competencia&"', @dt_i = '"&mes_ant&"/21/"&ano_ant&"', @dt_f = '"&dt_mes&"/20/"&dt_ano&"', @outras_cond='"&outras_cond&"', @nm_ordem='"&nm_ordem&"'"
					'xsql = "up_funcionario_rcm_geral @dt_comp_i='8/21/2012', @dt_comp_f='9/20/2012', @dt_i = '"&mes_ant&"/21/"&ano_ant&"', @dt_f = '"&dt_mes&"/20/"&dt_ano&"', @outras_cond='"&outras_cond&"', @nm_ordem='"&nm_ordem&"'"
						Set rs = dbconn.execute(xsql)
							while not rs.EOF
							
								strcod = rs("cd_codigo")
								cd_matricula = rs("cd_matricula")
								nm_nome = rs("nm_nome")
								cd_unidade = rs("cd_unidade")
								nm_unidade = rs("nm_unidade")
								nm_sigla = rs("nm_sigla")
								cd_hospital = rs("cd_hospital")
								cd_sexo = rs("cd_sexo")
								
								dt_admissao = rs("dt_admissao")
								dt_demissao = rs("dt_demissao")
								
								admissao = rs("admissao")
								demissao = rs("demissao")
									if demissao <> "null" then
										cor_registro = "red"
									else
										cor_registro = "black"
										demissao = ""
									end if
								
								'***** VT Cancelado **********************************************
								'xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '04/26/2012') AND (nr_valor > 0)"
								xsql = "SELECT * FROM TBL_Funcionario_valores WHERE cd_funcionario = "&strcod&" AND (cd_tipo = 4) AND (dt_atualizacao <= '"&dt_mes&"/"&dt_dia&"/"&dt_ano&"') AND (nr_valor > 0)"
								Set rs_vt_canc = dbconn.execute(xsql)
									if not rs_vt_canc.EOF Then
										cd_vt_canc = rs_vt_canc("nr_valor")
									end if
								'*****************************************************************
								
								'if admissao = year(now)&zero(month(now)) then
								''	dias_trab_admissao = datediff(m,dt_admissao,now)
								'end if
								'	
								'	'*** Caso a admissão for no mesmo mês de competencia da folha ***
								'	if admissao = dt_ano&zero(dt_mes)  AND demissao = "" OR admissao = dt_ano&zero(dt_mes) AND demissao > dt_ano&zero(dt_mes) then
								'		tempo_trabalhado = DateDiff("d",day(dt_admissao)&"/"&month(dt_admissao)&"/"&year(dt_admissao),dia_final&"/"&dt_mes&"/"&dt_ano)+1 '(1 = conta o dia da contratação)
								'	
								'	'*** Caso a admissão e a demissão forem no mesmo mês de competencia da folha ***
								'	elseif admissao = dt_ano&zero(dt_mes)  AND demissao = dt_ano&zero(dt_mes) then
								'		tempo_trabalhado = DateDiff("d",day(dt_admissao)&"/"&month(dt_admissao)&"/"&year(dt_admissao),dia_final&"/"&dt_mes&"/"&dt_ano)+1 '(1 = conta o dia da contratação)
								'	
								'	else
								'		tempo_trabalhado = "30"
								'	end if
								'
								expediente = rs("expediente")
									expediente = expediente' / 60
									t_expediente = expediente / 60
										if t_expediente >= 9 AND t_expediente <= 11 then
											tipo_expediente = "8h"
											carga_mensal = 220
										elseif t_expediente = 6 then
											tipo_expediente = "6h"
											carga_mensal = 180
										elseif t_expediente = 12 then
											tipo_expediente = "12x36h"
											carga_mensal = 180
										end if
								
								
									if num_cor = 1 then
										cor_linha = "#ccffff"
									else
										cor_linha = "#ffffff"
									end if
									
									'***************
									'*** Filtros ***
									'***************
									if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
										cond_1 = 1
										subtotal = 1
									end if
									
									if mostra_faltas <> 0 then 
										if int(mostra_faltas) = "1" AND  int(total_faltas) = 0 then
											cond_2 = 1
										elseif int(mostra_faltas) = "2" AND int(total_faltas) > 0 then
											cond_2 = 1
										end if
									else
										cond_2 = 1
									end if
									
								'**********
							if ultima_unidade =	"" OR ultima_unidade <> cd_unidade AND subtotal = 1 then
								if ultima_unidade <> "" Then%>
								<tr>
									<td colspan="2" align="center">&nbsp;<%'=strcd_unidade%></td>
									<td>Subtotal</td>
									<td>&nbsp;</td>
									<!--td align="center"><b><%=subtotal_he_limit%></b></td>
									<td align="center"><b><%=subtotal_he_plus%></b></td-->
									<td align="center"><b><%=total_mes(subtotal_he)%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_ad_noturno%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_vr_extra%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_vt_extra%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_falta_justif%></b></td>
									<td align="center"><b><%=subtotal_falta_injust%></b></td>
									<td align="center"><b><%=subtotal_faltas%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_atraso%></b></td>
								</tr>
								<tr><td colspan="17">&nbsp;</td></tr>
								<%'*** Limpa os subtotais ***
									subtotal_he_limit = 0
									subtotal_he_plus = 0
									subtotal_he = 0
									subtotal_ad_noturno = 0
									subtotal_vr_extra = 0
									subtotal_vt_extra = 0
									subtotal_falta_justif = 0
									subtotal_falta_injust = 0
									subtotal_faltas = 0
									subtotal_atraso = 0
										
									subtotal = 0
								end if
							end if		
						
							if strcd_unidade = 0 OR int(strcd_unidade) = cd_unidade then
								'***** Dados do RCM **********	
								'xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_competencia&"/1/"&ano_competencia&"', @dt_f='"&mes_competencia&"/"&final_competencia&"/"&ano_competencia&"'"
								xsql = "up_Funcionario_rcm_lista @cd_funcionario='"&strcod&"', @dt_i='"&mes_ant&"/21/"&ano_ant&"', @dt_f='"&mes_competencia&"/20/"&ano_competencia&"'"
								Set rs_rcm = dbconn.execute(xsql)
								if not rs_rcm.EOF then
									total_he = rs_rcm("total_he")
									total_ad_noturno = rs_rcm("total_ad_noturno")
									total_vr_extra = rs_rcm("total_vr_extra")
									total_vt_extra = rs_rcm("total_vt_extra")
									total_falta_justif = rs_rcm("total_falta_justif")
									total_falta_injust = rs_rcm("total_falta_injust")
									total_faltas = rs_rcm("total_faltas")
									total_atraso = rs_rcm("total_atraso")
										
										subtotal_ad_noturno = subtotal_ad_noturno + total_ad_noturno
										subtotal_vr_extra = subtotal_vr_extra + total_vr_extra
										subtotal_vt_extra = subtotal_vt_extra + total_vt_extra
										subtotal_falta_justif = subtotal_falta_justif + total_falta_justif
										subtotal_falta_injust = subtotal_falta_injust + total_falta_injust
										subtotal_faltas = subtotal_faltas + total_faltas
										subtotal_atraso = subtotal_atraso + total_atraso
											
											total_geral_ad_noturno = total_geral_ad_noturno + total_ad_noturno
											total_geral_vr_extra = total_geral_vr_extra + total_vr_extra
											total_geral_vt_extra = total_geral_vt_extra + total_vt_extra
											total_geral_falta_justif = total_geral_falta_justif + total_falta_justif
											total_geral_falta_injust = total_geral_falta_injust + total_falta_injust
											total_geral_faltas = total_geral_faltas + total_faltas
											total_geral_atraso = total_geral_atraso + total_atraso
								end if
								
									if isnumeric(total_he) then
										'total_he = (total_he/60)
										total_he = (total_he)
										'********************************************
										'*** Arredondamento Hora Extra fracionada ***
										'		if instr(1,total_he,",",1) > 0 then 
										'			virgula_he = instr(1,total_he,",",1)
										'			fracao_he = mid(total_he,virgula_he+1,len(total_he))
											'			if fracao_he > 2 then
										'					total_he = int(total_he) + 1
										'					sinal_arredonda = "*"
										'				else
										'					total_he = total_he
										'					sinal_arredonda = ""
										'				end if
										'		end if
										
										
										
											subtotal_he = subtotal_he + total_he
												total_geral_he = total_geral_he + total_he
										if total_he > 36 then
											total_he_plus = total_he-36
											total_he_limit = 36
												subtotal_he_limit = subtotal_he_limit + total_he_limit
												subtotal_he_plus = subtotal_he_plus + total_he_plus
													total_geral_he_limit = total_geral_he_limit + total_he_limit
													total_geral_he_plus = total_geral_he_plus + total_he_plus
										else
											total_he_plus = 0
											total_he_limit = total_he
												subtotal_he_limit = subtotal_he_limit + total_he_limit
												subtotal_he_plus = subtotal_he_plus + total_he_plus
													total_geral_he_limit = total_geral_he_limit + total_he_limit
													total_geral_he_plus = total_geral_he_plus + total_he_plus
										end if
									else
											total_he = 0
											total_he_plus = 0
											total_he_limit = 0
												subtotal_he_limit = subtotal_he_limit
												subtotal_he_plus = subtotal_he_plus
													total_geral_he_limit = total_geral_he_limit
													total_geral_he_plus = total_geral_he_plus
									end if
							end if%>
							
					<%'if int(strcd_unidade) = "0" OR int(cd_unidade) = int(strcd_unidade) then
					if cond_1 = 1 AND cond_2 = 1 then 
					'if int(filtros)="00" AND  int(total_faltas)>=0 AND int(cd_unidade) = int(strcd_unidade) OR int(filtros)="00" AND int(total_faltas)<=0 AND int(cd_unidade) = int(strcd_unidade) OR int(filtros)="00" AND  int(total_faltas)>=1 AND int(cd_unidade) = int(strcd_unidade)then
							
							
						if ultima_unidade =	"" OR ultima_unidade <> cd_unidade then%>
							<tr><td align="center" colspan="17" style="background-color:gray; color:white;"><b><%=nm_unidade%></b></td></tr></tr>
							<tr style="background-color:gray; color:white;">
								<td>&nbsp;<img src="../../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
								<td align="center"><b>Mat.</b><br><img src="../../imagens/px.gif" alt="" width="25" height="1" border="0"></td>
								<td align="center"><b>Funcion&aacute;rio</b><br><img src="../../imagens/px.gif" alt="" width="230" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<!--td align="center"><b>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>H.E. +</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td-->
								<td align="center"><!--b>Total<br--><b>H.E.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>ad.<br>Noturno</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>VR Extra</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>VT Extra</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
									<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center"><b>Faltas just.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Faltas injust.</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td align="center"><b>Total faltas</b><br><img src="../../imagens/px.gif" alt="" width="35" height="1" border="0"></td>
								<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<%'if strtipo <> "var_enf" then%><!--td align="center"><b>DSR.</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td--><%'end if%>
								<td align="center"><b>Atrasos</b><br><img src="../../imagens/px.gif" alt="" width="40" height="1" border="0"></td>
							</tr>							
						<%subtotal = 1
						ultima_unidade = cd_unidade
						end if%>
						
						<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=cor_linha%>');" bgcolor="<%=cor_linha%>">	
						<%variaveis1 = "empresa/funcionario/funcionarios_cadastro.asp?tipo=cadastro&cod="&strcod&"','cadastro','width=790,height=350,scrollbars=1,resizable=1"
						variaveis2 = "empresa/funcionario/funcionarios_folhapagamento_rcm_cad.asp?tipo=cadastro&cod="&strcod&"&dt_dia="&dt_dia&"&dt_mes="&dt_mes&"&dt_ano="&dt_ano&"&dia_sel="&dia_sel&"&mes_sel="&mes_sel&"&ano_sel="&ano_sel&"&var=1&cat="&strcat&"','variaveis','width=555,height=510,scrollbars=1"%>
													
								<td align="right" valign="top"><b><%=zero(num_funcionario)%></b>&nbsp;<%'=" - "&periodo_ferias%></td>
								<td align="right" valign="top" onClick="window.open('<%=variaveis1%>')"><%=cd_matricula%>&nbsp; </td>
								<td style="color:<%=cor_registro%>;" align="left" onClick="window.open('<%=variaveis1%>')"><%=nm_nome%> <%'if cd_vt_canc > 0 then response.write("*")%><!-- &nbsp; --><%'if expediente = 6 then response.write("**")%></td>
								<!--td style="color:<%=cor_registro%>;" align="center" onClick="window.open('<%=variaveis1%>')"><%=nm_sigla%></td-->
									
					<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Qtd Hora Extra *** -->
								<!--td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_he_limit > 0 then%><%=sinal_arredonda%><%=(total_he_limit)%><%else%>0<%end if%><%'=subtotal_he_limit%></td-->
					<!-- *** Qtd Hora Extra Excedentes*** -->
								<!--td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_he_plus > 0 then%><span style="color:<%if total_he_plus > 0 then%>#336600;<%else%>black;<%end if%>"><%=total_he_plus%><%else%>0</span><%end if%></td-->
					<!-- *** Total Horas Extras *** -->
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_he > 0 then%><%=formatahora(total_he)%><%else%>0<%end if%></td>
					<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					
					<!-- *** Qtd. Adicional Noturno *** -->
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')">&nbsp;<%if total_ad_noturno > 0 then%><%=int(total_ad_noturno)%><%else%>0<%end if%><!-- <%'if qtd_dia_noturno > "" then%> dias<%'elseif qtd_dia_noturno <> "" then%>dia<%'end if%>--></td>
					<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Benefícios *** -->			
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_vr_extra >= 1 then%><%=(total_vr_extra)%><%else%>0<%end if%></td>
					<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_vt_extra >= 1 then%><%=(total_vt_extra)%><%else%>0<%end if%></td>
						
					<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<!-- *** Faltas justificadas *** -->
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_falta_justif > 0 then%><%=(total_falta_justif)%><%else%>0<%end if%></td>
					<!-- *** Faltas Não justificadas *** -->
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_falta_injust > 0 then%><%=(total_falta_injust)%><%else%>0<%end if%></td>
					<!-- *** Total Faltas  *** -->
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_faltas > 0 then%><%=(total_faltas)%><%else%>0<%end if%></td>
					<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
					<%'if strtipo <> "var_enf" then%>
					<!-- *** DSR Faltas Não justificadas *** -->
								<!--td align="right"><%'if nr_dsr_faltas > 0 then%><%'=(nr_dsr_faltas)%><%'end if%></td><%'end if%>
					<!-- *** Atrasos *** -->
								<td align="center" valign="top" onClick="window.open('<%=variaveis2%>')"><%if total_atraso > 0 then%><%=total_atraso%><%else%>0<%end if%></td>
					<%'if strtipo <> "var_enf" then%><%'end if%>
					
							
						<%num_funcionario = num_funcionario + 1
							num_cor =  num_cor + 1
								if num_cor > 2 then
									num_cor = 1
								end if
						end if%>
								<%'*** H E ***
								total_he_limit = 0
								total_he = 0
								total_he_plus = 0
								'***********
								total_ad_noturno = 0
								'***********
								total_vr_extra = 0
								total_vt_extra = 0
								'***********
								total_falta_justif = 0
								total_falta_injust = 0
								total_faltas = 0
								'***********
								total_atraso = 0
								'***********
								sinal_arredonda = ""
								
								
								strcod = 0
								cd_matricula = 0
								nm_nome = 0
								tempo_trabalhado = ""
								cd_unidade = 0
								nm_unidade = 0
								nm_sigla = 0
								cd_hospital = 0
								cd_salario = 0
								nr_salario = 0
								nr_salario_hora = 0
								nr_salario_dia = 0
								nr_salario_desc = 0
								nr_desc_salario = 0
								insalubridade = 0
								nr_insalubridade_forca = 0
								total_faltas_mes = 0
								expediente = 0
								tipo_expediente = 0
								carga_mensal = 0
								cd_aj_custo = 0
								nr_aj_custo = 0
								cd_ad_noturno = 0
								qtd_dia_noturno = 0
								ad_noturno = 0
								dsr_ad_not = 0
								
								
								
								
								cond_1 = 0
								cond_2 = 0
								
								'if ultima_unidade =	"" OR ultima_unidade <> cd_unidade then
							'if ultima_unidade <> "" Then%>
								
								
								<%rs.movenext
								wend
								
								if subtotal = 1 then%>
								<tr>
									<td colspan="2" align="center">&nbsp;<%'=strcd_unidade%></td>
									<td>Subtotal</td>
									<td>&nbsp;</td>
									<!--td align="center"><b><%=subtotal_he_limit%></b></td>
									<td align="center"><b><%=subtotal_he_plus%></b></td-->
									<td align="center"><b><%=total_mes(subtotal_he)%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_ad_noturno%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_vr_extra%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_vt_extra%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_falta_justif%></b></td>
									<td align="center"><b><%=subtotal_falta_injust%></b></td>
									<td align="center"><b><%=subtotal_faltas%></b></td>
									<td>&nbsp;</td>
									<td align="center"><b><%=subtotal_atraso%></b></td>
								</tr>
								<tr><td colspan="19">&nbsp;</td></tr>
								<%end if
								
								'*** Limpa os subtotais ***
									subtotal_he_limit = 0
									subtotal_he_plus = 0
									subtotal_he = 0
									subtotal_ad_noturno = 0
									subtotal_vr_extra = 0
									subtotal_vt_extra = 0
									subtotal_falta_justif = 0
									subtotal_falta_injust = 0
									subtotal_faltas = 0
									subtotal_atraso = 0%>
						<tr><td colspan="19" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
						<tr>
								<!--td align="right" colspan="3">&nbsp;</td-->
								<td align="center" colspan="3">&nbsp;<b>TOTAL</b></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<!--td align="center">&nbsp;<b><%=total_geral_he_limit%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_he_plus%></b></td-->
								<td align="center">&nbsp;<b><%=total_mes(total_geral_he)%></b></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_ad_noturno%></b></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_vr_extra%></b></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_vt_extra%></b></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_justif%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_falta_injust%></b></td>
								<td align="center">&nbsp;<b><%=total_geral_faltas%></b></td>
							<td bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td>
								<td align="center">&nbsp;<b><%=total_geral_atraso%></b></td>								
						</tr>
						
										
						<tr><td colspan="19" bgcolor="silver"><img src="../../imagens/px.gif" alt="" width="1" height="1" border="0"></td></tr>
						<tr><td colspan="19">&nbsp;</td></tr>
						<!--tr><td colspan="19">&nbsp;<span style="font-size:9px;">* optou por não receber vale transporte &nbsp; / &nbsp; ** expediente de 6h não recebe vale refeição</span></td></tr-->
					</table>
		<%'end if%>