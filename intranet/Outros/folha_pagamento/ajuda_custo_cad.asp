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
						<td colspan="100%">&nbsp;<b>CMI - <font color="red">Ajuda de custo</font></b></td>
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
					<form action="../../adm_ajuda_custo.asp" name="Ano" method="get">
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
					<form action="../../adm_ajuda_custo.asp" name="mes" method="get">
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
								<%do while not mes_inicio = mes_prod + 2
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
			<table cellspacing="1" cellpadding="1" border="1"  class="table">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="5">[<a href="adm_plantao_noite.asp">Mudar Ano</a>] &nbsp;&nbsp; [<a href="adm_plantao_noite.asp?ano_sel=<%=ano_sel%>">Mudar Mês</a>] &nbsp;&nbsp;<b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b><%=encerra%></td>
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
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><b>&nbsp;</b></td>
						<td><b>Nome</b></td>
					</tr>
					<form action="adm/adm3_acao.asp" method="get">						
					<%numero = "1"
					num = numero
					
					'***********************************************************************************
					'****** Lista os funcionarios de acordo com a data de admissão e desligamento ******
					'***********************************************************************************
					admissao = "year(dt_contratado) <= "&ano_sel&" AND Month(dt_contratado) <= "&mes_sel&""
					demissao = "month(dt_desliga) >= "&mes_sel&""
					demissao_null = "dt_desliga is Null "
					
					'condicao = "where cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao_null&" OR cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao&" "
					condicao = "where cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao&" OR cd_regime_trabalho = 1 AND "&admissao&" AND "&demissao_null&""
					xsql = "SELECT * FROM VIEW_funcionarios_lista "&condicao&" Order by nm_nome, nm_sobrenome"
					Set rs = dbconn.execute(xsql)
					
					Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					nm_foto = rs("nm_foto")
					cd_unidades = rs("cd_unidades")
					dt_contratado = rs("dt_contratado")
					dt_desliga = rs("dt_desliga")
					cd_status = rs("cd_status")
					cd_regime_trabalho = rs("cd_regime_trabalho")
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
					
					<%if num = 1 then%>
					<tr>
					<%end if%>
						
						
						<td><%=numero%><%'=num%></td>
						<td><input type="checkbox" name="cd_funcionario" value="<%=cd_codigo%>"><b><%'=cd_codigo%></b></td>
						<td><%=nm_nome%>&nbsp;<%=nm_sobrenome%><%'=cd_funcao%></td>
						
					<%
					
					'numero = numero + 1
					'rs.movenext
					
					'loop%>
					
					
										
					
						
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
						num = num + 1
						if num > 2 then
							num = 1
						end if
					rs.movenext
					
					loop%>
						<td><%=total_funcionario%></td>
					</tr>
					<tr>
						<td colspan="100%">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="100%">
						<input type="hidden" name="cd_tipo" value="2">
						<input type="hidden" name="cd_mes" value="<%=mes_sel%>">
						<input type="hidden" name="cd_ano" value="<%=ano_sel%>">
						<input type="hidden" name="acao" value="inserir">
						<input type="hidden" name="voltar" value="../adm_aux_creche.asp?ano_sel=<%=ano_sel%>&mes_sel=<%=mes_sel%>">
						<input type="submit">
						<input type="reset"></td>
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
				</form>
				<tr><td>&nbsp;<br><br><br></td></tr>	
			</table>
			<%end if%>
		</td>
	</tr>
</table>
</body>
</html>
