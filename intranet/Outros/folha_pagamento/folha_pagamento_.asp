<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--include file="../../includes/util.asp"-->
<%tela = request("tela")

if tela = "0" or tela = "" Then
	div_height = "200px"
	tela = "0"
else
	tela = "1"
end if
%>
<style type="text/css">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_ {color: #000000;font-family: verdana;font-size:11px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:10px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #000000; }  
.coluna {background-color: #99ff99;}

.divrolagem {
/* define barra de rolagem automatica quando o 
conteudo ultrapassar o limite em x ou y */
    overflow: auto; 
/* define o limite maximo da autura do div */
    height: <%=div_height%>;
/* define o limite maximo da largura do div */
    width: auto;}
#no_print{ 
	visibility:visible; 
	display: block;}
#ok_print{
	visibility:hidden; 
	display: none;}
.table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 100%;
	font-family: verdana;
	font-size:11px;
	text-decoration:none;}
-->
</style>

<style type="text/css" media="print">
<!--
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
#no_print{ 
	visibility:hidden; 
	display: none;}
#ok_print{
	visibility:visible;
	display:block;}
table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 200px;
	font-family: verdana;
	font-size:10px;
	text-decoration:none;}
.divrolagem {
	overflow: auto; 
    height: auto;
    width: auto;}
-->
</style>
<style type="text/css" media="print">
<!--
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
#no_print{ 
	visibility:hidden; 
	display: none;}
#ok_print{
	visibility:visible;
	display:block;}
table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 200px;
	font-family: verdana;
	font-size:10px;
	text-decoration:none;}
.divrolagem {
	overflow: auto; 
    height: auto;
    width: auto;}
-->
</style>

<%
dt_inicio_sys = "01/03/2008"

dia_prod = day(now)
mes_prod = month(now)
ano_prod = year(now)

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

	if mes_sel < 10 Then
	'mes_sel = 0 & mes_sel
	mes_selec = 0 & mes_sel
	End if

'data_atual = ano_sel&mes_selec
data_atual = dia_sel&"/"&mes_sel&"/"&ano_sel
data_prod = dia_prod&"/"&mes_prod&"/"&ano_prod

if cd_origem = "" Then
	cd_codigo = request("cd_codigo")
End if



%>
<html>
<head>
	<TITLE> Folha de Pagamento </TITLE>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
</head>

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
document.form.diaentrada.focus(); }


  function mOvr(src,clrOver) {
    if (!src.contains(event.fromElement)) {
	  src.style.cursor = 'hand';
	  src.bgColor = clrOver;
	}
  }
  function mOut(src,clrIn) {
	if (!src.contains(event.toElement)) {
	  src.style.cursor = 'default';
	  src.bgColor = clrIn;
	}
  }

  
  
  
// -->	
</script>
<body>
<!--body onload="foco()"--><br>
<!--table cellspacing="0" cellpadding="0" align="center" border="1" width="100%" id="no_print">
	<tr>
		<td align="left"-->
			<table cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" class="table" width="100%" id="print">
					<tr bgcolor=#cococo id="no_print"><!--Separador e limitador de espaços-->
						<td><img src="imagens/px.gif" width=200 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
						<td><img src="imagens/px.gif" width=20 height=0></td>
					</tr>
					<tr>
						<td colspan="100%">&nbsp;<b>CMI - <font color="red">Folha de Pagamento - Planilha</font></b></td>
					</tr>
					<!--tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/blackdot.gif" width="100%" height=0></td>
					</tr-->
					
				<%'*********************************************************
				  '*				      1ª Parte	 					  '*
				  '*********************************************************
				  '* 1.1 - Seleção do ANO *
				  '************************
				  
						if year(dt_inicio_sys) = int(ano_prod) then
							'ano_sel = year(dt_inicio_sys)
						end if
					if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr><td align=center colspan="100%">&nbsp;</td></tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
					    <td colspan="10" align="center">Selecione o ano abaixo:</td>
					</tr>
					<form action="../../folha_pagamento.asp" name="Ano" method="get">
					<tr>
						<td align="center" colspan="100%"><br>
						<select name="ano_sel">
								<%ano_inicio = year(dt_inicio_sys)
								response.write(ano_inicio)
								do while not ano_inicio = ano_prod + 1%>
								<option value="<%=ano_inicio%>" <%if ano_inicio=ano_prod then%><%response.write("selected")%><%end if%>><%=ano_inicio%></option>
								<%ano_inicio = ano_inicio + 1
								loop%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>
					<%
					'************************
					'* 1.2 - Seleção do MÊS *
					'************************
					
					elseif mes_sel = "" AND ano_sel <> "" AND cd_codigo = "" Then%>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=1></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="5">[<a href="folha_pagamento.asp">Mudar Ano</a>] &nbsp;&nbsp; <%=seta_a1%><b><%=ano_sel%></b><%=seta_a1%></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td colspan="4"><br>&nbsp;</td>
					</tr>
					<form action="../../folha_pagamento.asp" name="mes" method="get">
					<input type="hidden" name="ano_sel" value="<%=ano_sel%>">
					<tr>
					    <td colspan="10" align="center">Selecione o mês abaixo:</td>
					</tr>
					<tr>
						<td align="center" colspan="100%">
						<%ano_inicio = year(dt_inicio_sys)
						  mes_inicio = month(dt_inicio_sys)
								if int(ano_sel) > int(ano_inicio) then
									mes_inicio = "1"
								end if%>
								
						<select name="mes_sel">
								<%do while not mes_inicio = mes_prod + 1
								response.write(mes_inicio&"-")
								response.write(mes_prod&".")%>
								<option value="<%=mes_inicio%>" <%if mes_inicio=mes_prod then%><%response.write("selected")%><%end if%>><%=mes_selecionado(mes_inicio)%></option>
								<%
								
								
								mes_inicio = mes_inicio + 1
								loop%>			
						</select><br>
						<br>
						<input type="submit" value="OK">						
						</td>
					</tr>
					</form>
					
					<tr>
					    <td colspan="10">
							<p>&nbsp;</p>
							<p>&nbsp;</p></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
					</tr>
					
				<%'*********************************************************
				  '*				      2ª Parte	 					  '*
				  '*********************************************************
				  '* 2.1 - Corpo da Planilha *
				  '***************************
				   
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
			<table cellspacing="1" cellpadding="1" border="1"  class="txt_pequeno">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="5">[<a href="folha_pagamento.asp">Mudar Ano</a>] &nbsp;&nbsp; [<a href="folha_pagamento.asp?ano_sel=<%=ano_sel%>">Mudar Mês</a>] &nbsp;&nbsp;<b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b><%=encerra%></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif" height="0"></td>
					</tr>
					<tr bgcolor=#cococo id="print"><!--Separador e limitador de espaços-->
						<td><img src="imagens/blackdot.gif" width=18 height=0></td>
						<td><img src="imagens/blackdot.gif" width=30 height=0></td>
						<td><img src="imagens/blackdot.gif" width=250 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><b>Cód.</b></td>
						<td><b>Nome</b></td>
						<td class="coluna"><b>Salário</b></td>
						<td><b>Aj. custo</b></td>
						<td><b>Insalub.</b></td>
						<td><b>Ad. Notu</b></td>
						<td><b>Creche</b></td>
						<td><b>Qtd Extra</b></td>
						<td><b>H. Extra</b></td>
						<td><b>Alimentação</b></td>
						<td><b>Transporte</b></td>
						<td><b>Inss</b></td>
						<td><b>Saúde</b></td>
						<td><b>Cesta</b></td>
						<td><b>Sind.</b></td>
						<td><b>Total</b></td>
					</tr>
					
						
					<%numero = "1"
					
					'*** Salário Minimo ****
					sal_minimo = 415.00
					piso_salarial = 572
					dt_sal_minimo = "01/03/2008"
					
'************************* VALORES ADICIONAIS ***********************************************************************
'*		****** Lista os funcionários que receberão ajuda de custo (2) Para o mês selecionado *****
			xsql = "select * from TBL_funcionario_adicionais where dt_mes <= '"&mes_sel&"' AND dt_ano <= '"&ano_sel&"' AND cd_tipo = 2 order by cd_codigo Desc"
			Set rs = dbconn.execute(xsql)
			if not rs.EOF Then
			cd_func_ajuda = rs("cd_funcionarios")
			else 
			cd_func_ajuda = ""
			end if
		
'*		****** Lista os funcionários que receberão insalubridade (3) Para o mês selecionado *****
			xsql = "select * from TBL_funcionario_adicionais where dt_mes <= '"&mes_sel&"' AND dt_ano <= '"&ano_sel&"' AND cd_tipo = 3 order by cd_codigo Desc"
			Set rs = dbconn.execute(xsql)
			if not rs.EOF Then
			cd_func_insalub = rs("cd_funcionarios")
			else 
			cd_func_insalub = ""
			end if
		
'*		****** Lista os funcionários que receberão adicional noturno (4) Para o mês selecionado *****
			xsql = "select * from TBL_funcionario_adicionais where dt_mes <= '"&mes_sel&"' AND dt_ano <= '"&ano_sel&"' AND cd_tipo = 4 order by cd_codigo Desc"
			Set rs = dbconn.execute(xsql)
			if not rs.EOF Then
			cd_func_ad_not = rs("cd_funcionarios")
			else 
			cd_func_ad_not = ""
			end if
			
'*		****** Lista os funcionários que receberão auxílio creche (5) *****
			'xsql = "select * from TBL_funcionario_adicionais where dt_mes <= '"&mes_sel&"' AND dt_ano <= '"&ano_sel&"' AND cd_tipo = 5 order by cd_codigo Desc"
			'Set rs = dbconn.execute(xsql)
			'if not rs.EOF Then
			'cd_func_aux_creche = rs("cd_funcionarios")
			'else 
			'cd_func_aux_creche = ""
			'end if
			
'*		****** Lista os funcionários que receberão vale refeição (7) *****
			xsql = "select * from TBL_funcionario_adicionais where dt_mes <= '"&mes_sel&"' AND dt_ano <= '"&ano_sel&"' AND cd_tipo = 7 order by cd_codigo Desc"
			Set rs = dbconn.execute(xsql)
			if not rs.EOF Then
			cd_func_v_alim = rs("cd_funcionarios")
			else 
			cd_func_v_alim = ""
			end if
			
'*		****** Lista os funcionários que receberão vale transporte (8) *****
			xsql = "select * from TBL_funcionario_adicionais where dt_mes <= '"&mes_sel&"' AND dt_ano <= '"&ano_sel&"' AND cd_tipo = 8 order by cd_codigo Desc"
			Set rs = dbconn.execute(xsql)
			if not rs.EOF Then
			cd_func_v_transp = rs("cd_funcionarios")
			else 
			cd_func_v_transp = ""
			end if
'*********************************************************************************************************************

					'*********************************
					'****** Lista o funcionario ******
					'*********************************
					admissao = "year(dt_contratado) <= "&ano_sel&" AND Month(dt_contratado) <= "&mes_sel&""
					demissao = "month(dt_desliga) >= "&mes_sel&""
					demissao_null = "dt_desliga is Null "
					
					'condicao = "where cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao_null&" OR cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao&" "
					condicao = "where cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao&" OR cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao_null&""
					xsql = "SELECT * FROM VIEW_funcionario "&condicao&" Order by nm_nome, nm_sobrenome"
					Set rs = dbconn.execute(xsql)
					
					Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					'nm_foto = rs("nm_foto")
					cd_unidades = rs("cd_unidades")
					dt_contratado = rs("dt_contratado")
					dt_desliga = rs("dt_desliga")
					cd_status = rs("cd_status")
					cd_regime_trabalho = rs("cd_regime_trabalho")
					nm_creche = rs("nm_creche")
					'cd_funcao = rs("cd_funcao")
					'cd_carga_horaria = rs("cd_carga_horaria")
					'cd_carga_diaria = rs("cd_carga_diaria")
					
					'**************************************
					'*** carrega o cargo do funcionario ***
					'**************************************
						'strsql_funcao = "Select * From TBL_funcionario_funcao Where cd_funcionario = '"&cd_codigo&"' AND DAY(dt_atualizacao) <= '"&dia_sel&"' AND MONTH(dt_atualizacao) <= '"&mes_sel&"' AND YEAR(dt_atualizacao) <= '"&ano_sel&"' Order by cd_codigo Desc"
						strsql_funcao = "Select * From TBL_funcionario_funcao Where cd_funcionario = '"&cd_codigo&"' AND MONTH(dt_atualizacao) <= '"&mes_sel&"' AND YEAR(dt_atualizacao) <= '"&ano_sel&"' Order by cd_codigo Desc"
						Set rs_funcao = dbconn.execute(strsql_funcao)
						if not rs_funcao.EOF then
						cd_funcao = rs_funcao("cd_funcao")
						end if
						
					if Not funcionarios_ativos = "" Then
					virgula = ","
					End if
					funcionarios_ativos = funcionarios_ativos & virgula & cd_codigo
					%>
					
					<tr>
						<td><%=zero(numero)%></td>
						<td><b><%=cd_codigo%></b></td>
						<td><a href="empresa_funcionarios_cad.asp?pag=1&cod=<%=cd_codigo%>&tipo=nm_status&busca="><%=nm_nome%>&nbsp;<%=nm_sobrenome%><%'=cd_funcao%></a></td>
					<%
					
					'numero = numero + 1
					'rs.movenext
					
					'loop%>
					
					
										
					<%'*******************************************
					'*** Calcula o valor de acordo com o cargo ***
					'*********************************************
					xsql_1 = "select * from TBL_empresa_valores where cd_funcao = '"&cd_funcao&"' AND (MONTH(dt_data) <= '"&mes_sel&"') AND (YEAR(dt_data) <= '"&ano_sel&"') AND cd_tipo = 1 order by cd_codigo Desc"
						Set rs_1 = dbconn.execute(xsql_1)
						if not rs_1.EOF Then
						cd_salario = rs_1("cd_valor")
						else 
						cd_salario = "0"
						end if%>
						<td align="right" class="coluna">&nbsp;<%=cd_salario%></td>
						
					<%'********************
'ATENÇÃO!			'*** Ajuda de custo ***
					'**********************
					'Verifica quais funcionarios receberão
					cd_ajuda_custo = instr(1,cd_func_ajuda,cd_codigo,0)
					if cd_ajuda_custo <> 0 Then
						xsql_2 = "select * from TBL_empresa_valores where cd_funcao = '"&cd_codigo&"' AND (MONTH(dt_data) <= '"&mes_sel&"') AND (YEAR(dt_data) <= '"&ano_sel&"') AND cd_tipo = 2 order by cd_codigo Desc"
						Set rs_2 = dbconn.execute(xsql_2)
						if not rs_2.EOF Then
						cd_aj_custo = rs_2("cd_valor")
						end if
						cd_aj_custo = "X"
					Else
						cd_aj_custo = "&nbsp;"
					end if
					%><td align="right">&nbsp;<%=cd_aj_custo%></td>
						
					<%'*********************
					'*** Insalubridade *****
					'***********************
					cd_insalubridade = instr(1,cd_func_insalub,cd_codigo,0)
					if cd_insalubridade <> 0 Then
						'xsql_3 = "select * from TBL_empresa_valores where cd_funcao = '"&cd_codigo&"' AND (MONTH(dt_data) <= '"&mes_sel&"') AND (YEAR(dt_data) <= '"&ano_sel&"') AND cd_tipo = 3 order by cd_codigo Desc"
						'Set rs_3 = dbconn.execute(xsql_3)
						'if not rs_3.EOF Then
						cd_insalubr = sal_minimo * 0.10
						'end if
					Else
						cd_insalubr = "0"
					end if
						'cd_insalubr = sal_minimo * 0.10
					%><td align="right" class="coluna"><%=cd_insalubr%></td>
						
					<%'***********************
					'*** Adicional Noturno ***
					'*************************
					'formula: 40% do salário / 220 * 7 * 1
					'Verifica se o funcionario está apto a receber
						'cd_funcionarios
						cd_ad_noite = instr(1,cd_func_ad_not,cd_codigo,0)
						if cd_ad_noite <> 0 Then
						cd_ad_noturn = FormatNumber(cd_salario * 0.4 / 220 * 7 * 1 ,2)
						Else
						cd_ad_noturn = "&nbsp;"
						end if
						%>
						<td align="right" class="coluna">&nbsp;<%=cd_ad_noturn%></td>
						
					<%'********************
					'*** Auxilio creche ***
					'**********************
					'Verifica se o funcionário está apto a receber
					
						'cd_aux_creche = instr(1,cd_func_aux_creche,cd_codigo,0)
						'if cd_aux_creche <> 0 Then
						''cd_aux_creche = "X"
						'cd_aux_creche = FormatNumber(piso_salarial * 0.2 ,2)
						'Else
						'cd_aux_creche = "&nbsp;"
						'end if
						
						if nm_creche = 1 Then
						'cd_aux_creche = "X"
						cd_aux_creche = FormatNumber(piso_salarial * 0.2 ,2)
						Else
						cd_aux_creche = "&nbsp;"
						end if%>
						<td align="right" class="coluna">&nbsp;<%=cd_aux_creche%></td>
						
					<%'***********************
'ATENÇÃO!			'*** Qtd. Horas Extras ***
					'*************************
					'Verifica se o funcionário está apto a receber%>
						<td align="right">&nbsp;<%=cd_qtd_he%></td>
						
					<%'*****************************
					'*** Calculo de Horas Extras ***
					'*******************************
						cd_valor_he = cd_salario / 220 * 1.90
						if cd_qtd_he > "0" Then
							cd_valor_he = formatnumber(cd_valor_he * cd_qtd_he,2)
						else
							cd_valor_he = "0"
						end if
						%>
						<td align="right" class="coluna"><%=cd_valor_he%></td>
						
					<%'**********************
'ATENÇÃO!			'*** Vale refeição ***
					'************************
					'Verifica o valor a ser descontado do funcionario
					cd_v_alim = instr(1,cd_func_v_alim,cd_codigo,0)
						if cd_v_alim = 0 Then
						cd_v_alim = "X"
						Else
						cd_v_alim = "0"
						end if%>
						<td align="right"><%=cd_v_alim%></td>

					<%'*********************
					'*** Vale Transporte ***
					'***********************
					'Desconto de 6% sobre o valor do salário Bruto
					'Verifica o valor a ser descontado do funcionario
					cd_v_transp = instr(1,cd_func_v_transp,cd_codigo,0)
						if cd_v_transp = 0 Then
						'cd_v_transp = "X"
						cd_v_transp = formatnumber(cd_salario * 0.06 ,2)
						Else
						cd_v_transp = "0"
						end if%>
						<td align="right" class="coluna"><%=cd_v_transp%></td>
						
					<%'*********************
					'*** Calculo de INSS ***
					'***********************
						'INSS = (Sal + Insal + ad. not + HE) x %
						inss = cd_salario + cd_insalubr + cd_noturn '+ cd_valor_he
						if cd_salario <= 911.70 then
							inss_aliq = 0.08
						elseif cd_salario >= 911.71 AND cd_salario <= 1519.50 then
							inss_aliq = 0.09
						elseif cd_salario >= 1519.51 then
							inss_aliq = 0.11
						end if
						
						if inss <> "" Then
							inss = inss * inss_aliq
						end if%>
						<td align="right" class="coluna"><%=inss%></td>
						<%'*** Convenio Médico ***
						'Verifica quem terá desconto%>
						<td align="right"><%'=cd_valor%>11</td>
						
						<%'*** Cesta Basica ***%>						
						<td align="right"><%'=cd_valor%>12</td>
						
						<%'*** Contribuição sindical ***
						'REGRAS:
						'1º - Desconto para o funcionário ingressante;
						'2º - 1 vez ao ano %>
						<td align="right"><%'=cd_valor%>13</td>
						<td align="right"><%'=cd_valor%>0</td>
						
					<%'*** Armazena dados para não exibir multiplos ***
						funcionario = cd_funcionario
						'************************************************
						
						'rs_1.movenext
						'loop%>
						<%
						'*** Libera as variáveis *** 
						cd_funcao = ""
						cd_salario = ""
						cd_aj_custo = ""
						cd_insalubr = ""
						cd_ad_noturn = ""
						cd_aux_creche = ""
						cd_qtd_he = ""
						'****************************
						numero = numero + 1
					rs.movenext
					
					loop%>
						<td><%=total_funcionario%></td>
					</tr>

</td></tr>				
				<tr bgcolor=#FFFFFF id="no_print"><!--Separador e limitador de espaços-->
						<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blakdot.gif" width="100%" height="1"></td></tr>
				<tr bgcolor=#FFFFFF id="no_print"><!--Separador e limitador de espaços-->
					<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
				<tr><td>&nbsp;<br><br><br></td></tr>	
			</table>
			<%end if%>
		</td>
	</tr>
</table>
</body>
</html>
