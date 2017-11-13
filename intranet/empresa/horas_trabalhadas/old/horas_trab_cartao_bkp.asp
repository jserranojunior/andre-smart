<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/inc_open_connection.asp"-->
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

if cd_origem = "" Then
	cd_codigo = request("cd_codigo")
End if
acao = request("acao")
	if acao = "copiar" then	
		nm_acao = "inserir"
	elseif acao = "" Then
		nm_acao = "inserir"
	else
		nm_acao = acao
	end if
	'response.write(acao)

strcod = request("cod")
strhora = request("cd_hora")
stracao = request("acao")
strstatus = request("status")

strmes_sel = request("mes_sel")
strano_sel = request("ano_sel")

cd_unidades = request("cd_unidades")
cd_funcao = request("cd_funcao")

If stracao = "insert1" Then

dia_in = request("diaentrada")
mes_in = request("mes_in")
hora_in = request("horaentrada")

dia_out = request("diasaida")
mes_out = request("mes_out")
hora_out = request("horasaida")

Else

strdiaentrada = request("diaentrada")
strhoraentrada = request("horaentrada")
strdiasaida = request("diasaida")
strhorasaida = request("horasaida")

End If

strdia_in  = dia_in&"-"&mes_in&"-"&strano_sel
strdia_out = dia_out&"-"&mes_out&"-"&strano_sel

xsql_verif = "up_Verifica_fechamento_mes @cd_mes='"&strmes_sel&"', @cd_ano='"&strano_sel&"', @cd_origem=2"
Set rs_verif = dbconn.execute(xsql_verif) 'Verifica o fechamento do mes

Do While Not rs_verif.EOF

strclose = rs_verif("nm_fechado")

If strclose = int("1") Then
encerra = "(encerrado)"
fechado = "1"

End If
rs_verif.movenext

loop
%>
<html>
<head>
	<TITLE> Horas Trabalhadas </TITLE>
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
<body onload="foco()"><br>
<table cellspacing="0" cellpadding="0" align="center" border="1" width="100">
	<tr>
		<td align="left">
			<table cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" class="table">
					<tr bgcolor=#cococo id="no_print"><!--Separador e limitador de espaços-->
						<td align=center colspan="2"><img src="imagens/px.gif" width=200 height=0></td>
						<td><img src="imagens/px.gif" width=1 height=0></td>
						<td align=center><img src="imagens/px.gif" width=150 height=0></td>
						<td align=center><img src="imagens/px.gif" width=1 height=0></td>
						<td align=center colspan="2"><img src="imagens/px.gif" width=200 height=0></td>
						<td><img src="imagens/px.gif" width=1 height=0></td>
						<td><img src="imagens/px.gif" width=80 height=0></td>
						<td><img src="imagens/px.gif" width=75 height=0></td>
					</tr>
					<tr>
						<td colspan="100%">&nbsp;<b>FUNCIONÁRIOS - <font color="red">Horas trabalhadas - Digitação</font></b></td>
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
							ano_sel = year(dt_inicio_sys)
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
					<form action="../../empresa_funcionarios_horas.asp" name="Ano" method="get">
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
						<td colspan="5">[<a href="empresa_funcionarios_horas.asp">Mudar Ano</a>] &nbsp;&nbsp; <%=seta_a1%><b><%=ano_sel%></b><%=seta_a1%></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td colspan="4"><br>&nbsp;</td>
					</tr>
					<form action="../../empresa_funcionarios_horas.asp" name="mes" method="get">
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
								<option value="<%=mes_inicio%>" <%if mes_inicio=mes_prod - 1 then%><%response.write("selected")%><%end if%>><%=mes_selecionado(mes_inicio)%></option>
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
						<td><a href="empresa/cargas_horarias/cargas_horarias.asp" target="_blank">Carga horária</a></td>
					</tr>
					
				<%'*********************************************************
				  '*				      2ª Parte	 					  '*
				  '*********************************************************
				  '* 1.2 - Seleção do Funcionário *
				  '********************************
				   
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
			<table cellspacing="1" cellpadding="1" border="0"  class="table">
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="5">[<a href="empresa_funcionarios_horas.asp">Mudar Ano</a>] &nbsp;&nbsp; [<a href="empresa_funcionarios_horas.asp?ano_sel=<%=ano_sel%>">Mudar Mês</a>] &nbsp;&nbsp;<b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b><%=encerra%></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif" height="0"></td>
					</tr>
					<tr>
						<td colspan="4"><br>&nbsp;</td>
						<td align="Center" colspan="2"><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Carga horária</b></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><b>Código</b></td>
						<td><b>Matrícula</b></td>
						<td><b>Nome</b></td>
						<td align="right"><b>Mensal</b></td>
						<td align="right"><b>Diária</b></td>
					</tr>
					
					<%conta_coop = "1"
					'condicao = " Where  cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_demissao >= "&data_atual&" OR cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_demissao = null "
					'condicao = " Where  cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND (MONTH(dt_desliga) >= "&mes_sel&") AND (YEAR(dt_desliga) >= "&ano_sel&") OR cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= "&mes_sel&") AND (YEAR(dt_contratado) <= "&ano_sel&") AND dt_desliga = null "
					'condicao = " Where  cd_regime_trabalho = 1 AND dt_contratado <= 31/"&mes_sel&"/"&ano_sel&" "'AND dt_desliga >= 31/"&mes_sel&"/"&ano_sel&" "'OR cd_regime_trabalho = 1 AND dt_contratado <= "&mes_sel&"/31/"&ano_sel&" AND dt_desliga = null "
					condicao = " Where  cd_regime_trabalho = 1"' AND dt_contratado <= "&mes_sel&"/31/"&ano_sel&" "'AND dt_desliga >= 31/"&mes_sel&"/"&ano_sel&" "'OR cd_regime_trabalho = 1 AND dt_contratado <= "&mes_sel&"/31/"&ano_sel&" AND dt_desliga = null "
					'xsql = "up_funcionarios_lista @cd_codigo='', @condicao='"&condicao&"', @ordem= 'nm_nome, nm_sobrenome'"
					'*** WHERE     (cd_regime_trabalho = 1) AND (MONTH(dt_contratado) <= 1) AND (YEAR(dt_contratado) <= 2009) AND (MONTH(dt_desliga) >= 1) AND (YEAR(dt_desliga)                       >= 2009) OR                   cd_regime_trabalho = 1 AND (MONTH(dt_contratado) <= 01) AND (YEAR(dt_contratado) <= 2009) AND dt_desliga = NULL  ***'
					xsql = "up_funcionario_lista @cd_codigo='', @condicao='"&condicao&"', @ordem= 'nm_nome, nm_sobrenome'"
					Set rs = dbconn.execute(xsql)
					
					Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					cd_unidades = rs("cd_unidades")
					'dt_demissao = rs("dt_demissao")
					dt_desliga = rs("dt_desliga")
					cd_status = rs("cd_status")

					'if dt_demissao <= data_atual Then
					if dt_desliga = data_atual Then
					'if dt_desliga = mes_sel Then
					cor_font = "#afafaf"
					end if
					
					'Mostra a Carga Horária de cada Funcionário e avisa caso esteja faltando
					xsql_ch = "select * from TBL_carga_horaria Where cd_funcionario='"&cd_codigo&"' AND cd_mes='"&mes_sel&"' AND cd_ano='"&ano_sel&"'"
					Set rs_ch = dbconn.execute(xsql_ch)
					if not rs_ch.EOF Then
					cd_carga_horaria = rs_ch("cd_carga_horaria")
					cd_carga_diaria = rs_ch("cd_carga_diaria")
					end if
					%>
					<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa_funcionarios_horas.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&cd_codigo=<%=cd_codigo%>&tela=<%=tela%>');" onmouseout="mOut(this,'');">
						<td>&nbsp;</td>
						<td><font color="<%=cor_font%>"><%=cd_codigo%></font></td>
						<td><font color="<%=cor_font%>"><%=cd_matricula%></font></td>
						<td><font color="<%=cor_font%>"><%=nm_nome%>&nbsp;<%=nm_sobrenome%><%'=cd_status%></font></td>
						<td align="right"><font color="<%=cor_font%>"><%=total_mes(cd_carga_horaria)%></font></td>
						<td align="right"><font color="<%=cor_font%>"><%=total_mes(cd_carga_diaria)%></font></td>
					</tr>
					<%
					
					
					rs.movenext
					cor_font = ""
					Loop
					%>
					<tr><td>&nbsp;</td><td colspan=2>&nbsp;</td></tr>
					<tr align="left">
						<td>&nbsp;</td>
						<td colspan=2>&nbsp;</td></tr>
					<tr>
						<td colspan="100%"><a href="cargas_horarias.asp" target="_blank">Carga horária</a></td>
					</tr>
				<%'*********************************************************
				  '*				      3ª Parte	 					   *
				  '*********************************************************
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo <> "" Then%>
					
					<%'xsql = "up_funcionarios_lista @cd_codigo='"&cd_codigo&"', @ordem='nm_nome, nm_sobrenome', @condicao=''"
					xsql = "up_funcionario_lista @cd_codigo='"&cd_codigo&"', @ordem='nm_nome, nm_sobrenome', @condicao=''"
					Set rs = dbconn.execute(xsql)
					
					'Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					'nm_foto = rs("nm_foto")
					cd_unidades = rs("cd_unidades")
					dt_contratado = rs("dt_contratado")
					dt_desliga = rs("dt_desliga")
					cd_status = rs("cd_status")
					'>>>cd_funcao = rs("cd_funcao")
					'cd_carga_horaria = rs("cd_carga_horaria")
					'cd_carga_diaria = rs("cd_carga_diaria")
					
					xsql_ch = "SELECT * FROM TBL_carga_horaria WHERE cd_funcionario = '"&cd_codigo&"' AND cd_mes >= '"&mes_sel&"'  ORDER BY cd_mes,cd_ano DESC"
					set rs_ch = dbconn.execute(xsql_ch)
					
					if not rs_ch.EOF Then
					cd_carga_horaria = rs_ch("cd_carga_horaria")/60
					cd_carga_diaria = rs_ch("cd_carga_diaria")/60
					end if
					
					
					if cd_carga_horaria = "0" OR cd_carga_horaria = "" Then
					cd_carga_horaria = rs("cd_carga_horaria")
					Mens_CHM = (" *** ATENÇÂO: A carga horária não foi definida ou está definida como 0 (zero), a soma das horas será prejudicada *** ")
					end if
					
					if cd_carga_diaria = "" OR cd_carga_diaria = 0 Then
					cd_carga_diaria = rs("cd_carga_diaria")
					Mens_CHM = (" *** ATENÇÂO: A carga horária não foi definida ou está definida como 0 (zero), a soma das horas será prejudicada *** ")
					end if
					
					if dt_demissao <= data_atual Then
					status = "(desligado)"
					end if
					
					
					%>
					<tr bgcolor=#cococo id="no_print">
						<td align=center colspan="100%"><img src="imagens/px.gif" height=0></td>
					</tr>
					<tr id="no_print">
						<td></td>
						<td colspan="9">
						[<a href="empresa_funcionarios_horas.asp">Mudar Ano</a>] &nbsp;&nbsp; [<a href="empresa_funcionarios_horas.asp?ano_sel=<%=ano_sel%>">Mudar Mês</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp;[<a href="empresa_funcionarios_horas.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>">Mudar Funcionário</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp; [<a href="empresa_funcionarios_horas.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&cd_codigo=<%=cd_codigo%>&tela=1">Tela inteira</a>:<a href="empresa_funcionarios_horas.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&cd_codigo=<%=cd_codigo%>&tela=0">Tela normal</a>]</td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif" height=0></td>
					</tr>
					<tr bgcolor=#FFFFFF id="ok_print"><!-- espaçador impressão -->
						<td><img src="imagens/px.gif" width=28 height=0></td>
						<td><img src="imagens/px.gif" width=50 height=0></td>
						<td><img src="imagens/px.gif" width=80 height=0></td>
						<td><img src="imagens/px.gif" width=50 height=0></td>
						<td><img src="imagens/px.gif" width=60 height=0></td>
						<td><img src="imagens/px.gif" width=60 height=0></td>
						<td><img src="imagens/px.gif" width=130 height=0></td>
					</tr>
					<tr>
						<td align=left colspan="4" bgcolor="#C0C0C0"><br>&nbsp;<%=nm_nome%>&nbsp;<%=nm_sobrenome%> <%'=" - "&month(dt_desliga)&"/"%><%'=year(dt_desliga)%><b><br>&nbsp;</td>
						<%'Meses anteriores
						if mes_sel = 1 then 
							mes_ant = 12
							ano_ant = strano_sel - 1
						else
							mes_ant = mes_sel - 1
							ano_ant = strano_sel
						end if
							If month(dt_contratado) = int(mes_sel) AND year(dt_contratado) = int(strano_sel) then
								seta_a = ""
							Else
								seta_a = "<"
							end if
						
						'Meses posteriores
						if mes_sel = 12 then 
							mes_prox = 1
							ano_prox = strano_sel + 1
						else
							mes_prox = mes_sel +1
							ano_prox = strano_sel
						end if
							'If int(right(dt_demissao,2)) = int(mes_sel) AND int(left(dt_demissao,4)) = int(strano_sel) Or int(mes_sel) = int(mes_prod) then
							'If int(right(dt_desliga,2)) = int(mes_sel) AND int(left(dt_desliga,4)) = int(strano_sel) Or int(mes_sel) = int(mes_prod) then
							If int(month(dt_desliga)) = int(mes_sel) AND int(year(dt_desliga)) = int(strano_sel) Or int(mes_sel) = int(mes_prod) then
								seta_p = ""
							Else
								seta_p = ">"
							end if%>
							
			    		<td align=center colspan="4" bgcolor="#C0C0C0"><a href="empresa_funcionarios_horas.asp?mes_sel=<%=mes_ant%>&ano_sel=<%=ano_ant%>&cd_codigo=<%=cd_codigo%>&tela=<%=tela%>"><%=seta_a%></a>&nbsp;<%=mes_selecionado(mes_sel)%>/<%=strano_sel%></b>&nbsp;<a href="empresa_funcionarios_horas.asp?mes_sel=<%=mes_prox%>&ano_sel=<%=ano_prox%>&cd_codigo=<%=cd_codigo%>&tela=<%=tela%>""><%=seta_p%></a></td>
						<!--td align=center colspan="2"  bgcolor="#C0C0C0">
							<%' If nm_foto <> "" then%>
							<img src="foto/funcionarios/<%=nm_foto%>" width=100 >
							<%'Else%>
							<img src="imagens/Vdlap2p.gif" width=100>
							<br><br>
							<%'End if%>
						</td-->
			  		</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					
				</table>
				<table cellspacing="1" cellpadding="1" border="0" class="table" width=500>
				<tr bgcolor=#FFFFFF id="no_print">
						<!--td><img src="imagens/blackdot.gif" width=25 height=0--></td>
						<td><img src="imagens/px.gif" width=45 height=0></td>
						<td><img src="imagens/px.gif" width=75 height=0></td>
						<td><img src="imagens/px.gif" width=110 height=0></td>
						<td><img src="imagens/px.gif" width=110 height=0></td>
						<td><img src="imagens/px.gif" width=105 height=0></td>
						<td><img src="imagens/px.gif" width=110 height=0></td>
						<td><img src="imagens/px.gif" width=145 height=0></td>
						<td><img src="imagens/px.gif" width=65 height=0></td>
						<td><img src="imagens/px.gif" width=15 height=0></td>
					</tr>
				<tr><td colspan="100%"><font color = "red" id="no_print"><%=mens_CHM%></font></td></tr>
				<tr bgcolor="#cococo">
					<!--td>&nbsp;</td-->
					<td align=center>Dia</td>
					<td align=center>Entrada</td>
					<td align=center>Intervalo</td>
					<td align=center>Saida</td>
					<td align=center>Total</td>
					<td align=center>Acumulado<br>de horas</td>
					<td align="center">Observações</td>
					<td align=center id="no_print">Ações</td>
					<td id="no_print">&nbsp;</td>
				<tr>
				<tr>
					<td align=center colspan="100%" id="ok_print"><img src="imagens/blackdot.gif" width="100%" height=0></td>
				</tr>
				<tr bgcolor=#FFFFFF id="ok_print">
					<td><img src="imagens/px.gif" width=30 height=0></td>
					<td><img src="imagens/px.gif" width=50 height=0></td>
					<td><img src="imagens/px.gif" width=80 height=0></td>
					<td><img src="imagens/px.gif" width=50 height=0></td>
					<td><img src="imagens/px.gif" width=60 height=0></td>
					<td><img src="imagens/px.gif" width=60 height=0></td>
					<td><img src="imagens/px.gif" width=130 height=0></td>					
				</tr>
					<%
					primeiro_dia = 1
					dia_mes = primeiro_dia
					ultimo_dia = UltimoDiaMes(mes_sel,ano_sel)
					'ultimo_dia = 10
					'************************************************************
								'Verifica se há carga horaria para o mes de trabalho
							'xsql = "SELECT * From TBL_carga_horaria where cd_funcionario = '"&cd_codigo&"' AND cd_mes = '"&strmes_sel&"' AND cd_ano = '"&strano_sel&"'"
							'Set rs = dbconn.execute(xsql)
							'Do while not rs.EOF
							
							''cd_carga_horaria = rs("cd_codigo")
							'cd_carga_horaria = rs("cd_carga_horaria")
							'cd_carga_diaria = rs("cd_carga_diaria")
							'cd_carga_semana = cd_carga_diaria * 5
							''cd_carga_horaria = cd_carga_semana * 5
							''response.write("teste")
							''cd_carga_horaria = replace(cd_carga_horaria,":",",")
								''if cd_carga_horaria <> "" Then
									''cd_carga_horaria = cd_carga_horaria '* 60
									'cd_carga_horaria = cd_carga_horaria / 60
									'cd_carga_diaria = cd_carga_diaria / 60
									''cd_carga_diaria = cd_carga_diaria / 60
								''end if
							
							'rs.movenext
							'Loop
						'************************************************************%>
<%'************** Inicia uma tabela aqui *******************%>
						<tr>
							<td colspan="100%">
								<div class="divrolagem">
									<table cellspacing="1" cellpadding="1" border="0" class="table">
										<tr bgcolor=#FFFFFF id="no_print">
					<!--td><img src="imagens/px.gif" width=8 height=0></td-->
					<td><img src="imagens/px.gif" width=45 height=0></td>
					<td><img src="imagens/px.gif" width=75 height=0></td>
					<td><img src="imagens/px.gif" width=110 height=0></td>
					<td><img src="imagens/px.gif" width=110 height=0></td>
					<td><img src="imagens/px.gif" width=105 height=0></td>
					<td><img src="imagens/px.gif" width=110 height=0></td>
					<td><img src="imagens/px.gif" width=150 height=0></td>
					<td><img src="imagens/px.gif" width=70 height=0></td>
					<%if tela = "1" Then%><td>
						<img src="imagens/px.gif" width=13 height=0></td>
					<%end if%>
				</tr>
<%'*********************************************************
					conta_dias = 0
do while not dia_mes = ultimo_dia + 1

					str_dia = dia_mes&"/"&mes_sel&"/"&ano_sel
					
						nr_data = ano_sel&mes_sel&dia_mes
						nr_desliga = year(dt_desliga)&month(dt_desliga)&day(dt_desliga)
					
					dia_semana = weekday(str_dia)
						if dia_semana = 1 then 'or dia_semana = 7 then
							dia_cor = "red"
						else
							dia_cor = "black"
						end if
						
						if nr_data = nr_desliga then
							str_desligamento = "último dia"
						end if
					
					xsql2 = "up_Hora_Trabakhada_por_cd_funcionario_4 @cd_funcionario="&cd_codigo&",@dia="&dia_mes&",@mes="&mes_sel&",@ano='"&ano_sel&"'"
					Set rs2 = dbconn.execute(xsql2)
					
					do while not rs2.eof
						cd_folga = ""
						cd_cod_horas = rs2("cd_codigo")
						nm_obs = rs2("nm_obs")
						
						data_entrada_day=Day(rs2("data_entrada"))
						data_entrada_month=month(rs2("data_entrada"))
						data_entrada_year=year(rs2("data_entrada"))
						data_entrada_hour=hour(rs2("data_entrada"))
						data_entrada_minute=minute(rs2("data_entrada"))
						data_entrada_second=second(rs2("data_entrada"))
						data_entrada = data_entrada_month&"/"&data_entrada_day&"/"&data_entrada_year&" "&data_entrada_hour&":"&data_entrada_minute
						
						data_sai_day=day(rs2("data_saida_interv"))
						data_sai_month=month(rs2("data_saida_interv"))
						data_sai_year=year(rs2("data_saida_interv"))
						data_sai_hour=hour(rs2("data_saida_interv"))
						data_sai_minute=minute(rs2("data_saida_interv"))
						data_sai = data_sai_month&"/"&data_sai_day&"/"&data_sai_year&" "&data_sai_hour&":"&data_sai_minute
							
						data_entr_day=day(rs2("data_entr_interv"))
						data_entr_month=month(rs2("data_entr_interv"))
						data_entr_year=year(rs2("data_entr_interv"))
						data_entr_hour=hour(rs2("data_entr_interv"))
						data_entr_minute=minute(rs2("data_entr_interv"))
						data_entr = data_entr_month&"/"&data_entr_day&"/"&data_entr_year&" "&data_entr_hour&":"&data_entr_minute
						
						data_saida_day=Day(rs2("data_saida"))
						data_saida_month=month(rs2("data_saida"))
						data_saida_year=year(rs2("data_saida"))
						data_saida_hour=hour(rs2("data_saida"))
						data_saida_minute=minute(rs2("data_saida"))
						data_saida_second=second(rs2("data_saida"))
						data_saida = data_saida_month&"/"&data_saida_day&"/"&data_saida_year&" "&data_saida_hour&":"&data_saida_minute
					
						'*** Data de entrada Padrão brasileiro
						entrada_br = data_entrada_day&"/"&data_entrada_month&"/"&data_entrada_year&" "&data_entrada_hour&":"&data_entrada_minute
						
						'*** Procura dia duplicado no banco de dados ***
						
						str_duplos = instr(1,duplicados, data_entrada_day,1)
						duplicados = duplicados&","&data_entrada_day
						'***********************************************
					
						'*************** Calcula intervalo ***********************************************
						'Se saida ou entrada do intervalo for igual a nada, então calcular a diferença entre a hora de saída e a hora de entrada.
								if hour(data_sai)< hour(data_entrada) AND day(data_entrada)= ultimo_dia Then ' **** intervalo após a meia noite ***
									min1 = datediff("n",cdate(data_entrada),cdate(data_entrada_month&"/"&data_entrada_day&"/"&data_entrada_year&" "&"23:59"))
									min2 = datediff("n",cdate(data_entrada_month + 1 &"/"& 1  &"/"&data_entrada_year&" 0:0"),cdate(data_sai))
									min3 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = 1 + min1 + min2 + min3
									mens = "(a) "
								elseif hour(data_sai)< hour(data_entrada) Then ' **** intervalo após a meia noite ***
									min1 = datediff("n",cdate(data_entrada),cdate(data_entrada_month&"/"&data_entrada_day&"/"&data_entrada_year&" "&"23:59"))
									min2 = datediff("n",cdate(data_entrada_month&"/"&data_entrada_day +1 &"/"&data_entrada_year&" 0:0"),cdate(data_sai))
									min3 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = 1 + min1 + min2 + min3
									mens = "(1) "
								Elseif hour(data_entr)< hour(data_sai) Then ' ***** Intervalo em dias diferentes
									min0 = datediff("n",cdate(data_entrada),cdate(data_sai))
									min1 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = min0 + min1
									mens = "(2) "
								Elseif hour(data_saida)< hour(data_entrada) Then ' ***** Saída no outro dia
									min0 = datediff("n",cdate(data_entrada),cdate(data_sai))
									min1 = datediff("n",cdate(data_entr),cdate(data_entr_month&"/"&data_entr_day&"/"&data_entr_year&" 23:59"))
									min2 = datediff("n",cdate(data_saida_month&"/"&data_saida_day&"/"&data_saida_year&" 0:0"),cdate(data_saida))
									min = 1 + min0 + min1 + min2
									mens = "(3) "
								Else
									min1 = datediff("n",cdate(data_entrada),cdate(data_sai))
									min2 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = min1 + min2
									mens = "(0) "
								end if
								
								'Calcula o tempo de intervalo
								'tempo_interv = datediff("n",cdate(data_sai),cdate(data_entr))
								if hour(data_sai)> hour(data_entr) then 'Intervalo entre dois dias
								'tempo_interv = datediff("n",cdate(data_sai),cdate(data_entr))
								interv1 = datediff("n",cdate(data_sai),cdate(data_sai_month&"/"&data_sai_day&"/"&data_sai_year&" 23:59"))+1
								interv2 = datediff("n",cdate(data_entr_month&"/"&data_entr_day&"/"&data_entr_year&" 00:00"),cdate(data_entr))
								tempo_interv = interv1 + interv2
								'response.write("caso_1")
								else
								tempo_interv = datediff("n",cdate(data_sai),cdate(data_entr))
								'response.write("caso2")
								end if
								'response.write(cd_carga_horaria&" - ")
								'response.write(cd_carga_diaria)
								
								
								'******** Tempo de intervalo - Não desconta para quem faz 5,6 ou 12 horas *******
								If cd_carga_diaria = 5 or cd_carga_diaria = 6 OR cd_carga_diaria = 12 then
								intervalo_livre = tempo_interv
								min = min + intervalo_livre
								End if
								
								''******* Saldo das cargas horárias *******************
								saldo_dia = min - cd_carga_diaria * 60
								saldo_mes = saldo_mes + saldo_dia
									soma_diario = soma_diario + saldo_dia
									
						'***********************************************************
						'*	Verifica se há horas abonadas para os dias trabalhados *
						'***********************************************************
						'xsql3 = "up_funcionario_abono @cd_funcionario="&cd_codigo&", @dt_abono='"&mes_sel&"/"&dia_mes&"/"&ano_sel&"'"
						'Set rs3 = dbconn.execute(xsql3)
						''response.write(cd_codigo&"-")
						''response.write(dia_mes&"/"&mes_sel&"/"&ano_sel)
						
						'Do while not rs3.EOF
						'	dt_abono_tempo = rs3("dt_abono_tempo")
						'	saldo_dia = saldo_dia + dt_abono_tempo
						'	min = min + dt_abono_tempo
						'	msg_1 = "+"&total_mes(dt_abono_tempo)&"hs"
						'rs3.movenext
						'loop
									
									'*** Procura saldo negativo e avisa ***
									negativo_dia = Instr(1,saldo_dia,"-",1)'Compara se o caracter "-" aparece na string min
									If negativo_dia = "1" Then
									cor_saldo_dia = "red" 
									end if
									'****************************************************************************
									
							''******* Saldo das cargas horárias *******************
							'	saldo_dia = min - cd_carga_diaria
							'	saldo_mes = saldo_mes + saldo_dia
						'**********************************************************************************
						
						'*** Procura jornada de trabalho negativa e avisa ***
						negativo = Instr(1,min,"-",1)'Compara se o caracter "-" aparece na string min
						If negativo = "1" Then
						strst_ = "0" 
						strsinal = "ERRO!"
						end if
						'****************************************************************************
						'***********************************************************
												
						IF strst_ = "" Then
							strcor = "#eaeaea"
							str_cor_dia = "#eaeaea"
						ElseIf strst_ ="5"Then
							strcor = "#E0FFFF"
						Elseif strst_ = "0" Then
							strcor = "#ff8080"
						End If 
						
						if str_duplos <> 0 then
						str_cor_dia = "#FF0000"
						strcor = "#FF0000"
						str_aviso = "(Duplicado)"
						end if
					
						'*** Subtotal das horas ***
						strsun = strsun + min 
						
							'*** O que passar a  carga horária mensal será considerada hora extra ***
							'if strsun/60 > cd_carga_horaria then
								'response.write("***** igual ou maior! *****")
								'cd_carga = cd_carga_horaria * 60
								'saldo_dia = strsun - cd_carga
									'if saldo_dia > min then
									'response.write("é")
									'saldo_dia = min
									'end if
								'saldo_mes = saldo_mes + saldo_dia
								'strsun = strsun - saldo_dia
							'	'saldo_dia = strsun - cd_carga_horaria * 60 - saldo_dia
							'	saldo_mes = strsun - cd_carga_horaria * 60 - saldo_dia
							'end if
						
						'**************************
						'*** Dia da semana ***
						'dia_semana = weekday(entrada_br)
						'	if dia_semana = 1 or dia_semana = 7 then
						'		dia_cor = "red"
						'	else
						'		dia_cor = "black"
						'	end if
							
						conta_dias = conta_dias + 1
						
						'***********************************************************
						'*	Verifica se há horas abonadas para os dias trabalhados *
						'***********************************************************
						xsql3 = "up_funcionario_abono @cd_funcionario="&cd_codigo&", @dt_abono='"&mes_sel&"/"&dia_mes&"/"&ano_sel&"'"
						Set rs3 = dbconn.execute(xsql3)
						'response.write(cd_codigo&"-")
						'response.write(dia_mes&"/"&mes_sel&"/"&ano_sel)
						
						Do while not rs3.EOF
							cd_cod_abono = rs3("cd_codigo")
							dt_abono_tempo = rs3("dt_abono_tempo")
							saldo_dia = saldo_dia + dt_abono_tempo
							min = min + dt_abono_tempo
							msg_1 = "+"&total_mes(dt_abono_tempo)&"hs"
						rs3.movenext
						loop
						
'********************************************************
'** 1		/Linhas dos dias do mes trabalhado\		  ***
'********************************************************

%>
					<tr bgcolor="<%=str_cor_dia%>" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=strcor%>');" class="txt_">
						<!--td><font color="<%=dia_cor%>"><%=Left(DiaSemanaExtenso(dia_semana),3)%></font> </td-->
						<td align=center><%'=mens%><font color="<%=dia_cor%>"><%=data_entrada_day%></font></td>
						<td align=center><%=data_entrada_hour%>:<%=data_entrada_minute%>:<%=data_entrada_second%></td>
						<td align=center><%if data_sai = data_entrada Then%><%response.write("&nbsp;-")%><%else%><%=data_sai_hour%>:<%=data_sai_minute%> às <%=data_entr_Hour%>:<%=data_entr_minute%><%end if%></td>
						<td align=center><%=data_saida_hour%>:<%=data_saida_minute%>:<%=data_saida_second%></td>
						<td align=center><%if strsinal = "" Then%><%=total_mes(min)%><%else%><%=strsinal%><%end if%></td>
						<td align=center bgcolor="<%=cor_saldo_dia%>">&nbsp;<%=BH%><%=total_mes(saldo_dia)%></td>
						<td align=left>&nbsp;<%=str_aviso%> <%=nm_obs%> <%'=total_mes(tempo_interv)%> <%'=total_mes(tempo_interv2)%> <%'=carga_diaria%><%=msg_1%><!--(<%'=total_mes(strsun)%>) (<%'=cd_carga_horaria%>) (<%'=total_mes(saldo_dia)%>)--><%=str_desligamento%></td>
					<%If fechado = "1" Then%>
					    <Td align=center id="no_print"><input type="button" value="   ...   "></td>
					<%Else
					'cd_carga_horaria = cd_carga_horaria'*60
					excedente = strsun - cd_carga_horaria%>
	
						<td align="right" valign="middle" id="no_print">

							<!--input onclick="javascript:JsInsert('<%=data_entrada_day%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="I" title="Incluir"-->
							<img src="imagens/ic_copy.gif" onclick="javascript:JsCopy('<%=data_entrada_day%>','<%=data_entrada_month%>','<%=data_entrada_year%>','<%=cd_codigo%>','<%=tela%>')" type="button" value="C" title="copiar"> 						
							<a href="javascript:void(0)" onclick="window.open('empresa/abono/abono_ausencias.asp?cd_funcionario=<%=cd_codigo%>&dt_data=<%=str_dia%>', 'principal', 'width=200, height=420'); return false;"><img src="imagens/ic_clock.gif" border="0"></a>
							<img src="imagens/ic_editar.gif" onclick="javascript:JsEdit('<%=data_entrada_day%>','<%=data_entrada_month%>','<%=data_entrada_year%>','<%=cd_codigo%>','<%=tela%>')" type="button" value="E" title="Editar"> 
							<a href="javascript:void(0)" onclick="window.open('empresa/horas_trabalhadas/horas_trab_cartao_excluir.asp?cd_cod_horas=<%=cd_cod_horas%>&cd_cod_abono=<%=cd_cod_abono%>', 'principal', 'width=100, height=200'); return false;"><img src="imagens/ic_del.gif" type="button" value="A" title="Apagar" border="0"></a>
							<!--img src="imagens/ic_del.gif" onclick="javascript:JsDelete('<%=cd_codigo%>','<%=cd_cod_horas%>','<%=min%>','<%=min%>','<%=tela%>','<%=cd_cod_abono%>')" type="button" value="A" title="Apagar"-->
						</td>
							<%if tela = "1" Then%>
						<td id="no_print"><img src="../../imagens/px.gif" width=10 height=0></td>
							<%end if%>
					<%end if%>
					</tr>
					<%rs2.movenext
					str_aviso = ""
					cd_cod_horas = ""
					cd_cod_abono = ""
					loop


'***********************************************
'** 2		/Linhas dos dias de folga\		 ***
'***********************************************
'***********************************************************
						'*	Verifica se há horas abonadas para os dias trabalhados *
						'***********************************************************
						xsql3 = "up_funcionario_abono @cd_funcionario="&cd_codigo&", @dt_abono='"&mes_sel&"/"&dia_mes&"/"&ano_sel&"'"
						Set rs3 = dbconn.execute(xsql3)
						'response.write(cd_codigo&"-")
						'response.write(dia_mes&"/"&mes_sel&"/"&ano_sel)
						
						Do while not rs3.EOF
							cd_cod_abono = rs3("cd_codigo")
							dt_abono_tempo = rs3("dt_abono_tempo")
							saldo_dia = saldo_dia + dt_abono_tempo
							min = min + dt_abono_tempo
							msg_2 = "+"&total_mes(dt_abono_tempo)&"hs"
						rs3.movenext
						loop

	str_min = "&nbsp;"
	dt_data = mes_sel&"/"&dia_mes&"/"&ano_sel
	xsql3 = "Select * From TBL_hora_trabalhada_folga Where cd_funcionario='"&cd_codigo&"' AND dt_data='"&dt_data&"' ORDER BY dt_data"
		Set rs3 = dbconn.execute(xsql3)
			Do while NOT rs3.EOF
			cd_folga = rs3("cd_folga")
			dt_data = rs3("dt_data")
			txt_obs = rs3("txt_obs")
			cd_cod_folga = rs3("cd_codigo")
			
				if cd_folga = "2" Then
					nm_dsr = "(DSR)"
					'str_min = total_mes(cd_carga_diaria)
					str_min = total_mes(cd_carga_diaria*60)
					min = cd_carga_diaria*60
					strsun = strsun + min
				end if
			
				if cd_folga = 1 OR cd_folga = 2 then
					str_msg_ausencia = "&nbsp;&nbsp;&nbsp;&nbsp; F&nbsp;&nbsp;&nbsp; O&nbsp;&nbsp;&nbsp; L&nbsp;&nbsp;&nbsp; G&nbsp;&nbsp;&nbsp; A&nbsp;&nbsp;&nbsp;&nbsp;"
					str_cor_dia = "#bfffcf"
				elseif cd_folga = 3 then
					str_msg_ausencia = "&nbsp;&nbsp;&nbsp;&nbsp; F&nbsp;&nbsp;&nbsp; A&nbsp;&nbsp;&nbsp; L&nbsp;&nbsp;&nbsp; T&nbsp;&nbsp;&nbsp; A&nbsp;&nbsp;&nbsp;&nbsp;"
					str_cor_dia = "#ffffb9"
				end if
			
			'*** Procura dia duplicado no banco de dados ***
						
						'str_duplos_2 = instr(1,duplicados_2, day(dt_dia_mes),1)
						'duplicados_2 = duplicados_2&","&day(dt_dia_mes)
						'***********************************************
			
			
	'if data_entrada_day = "" AND dt_data <> "" then
		'if data_entrada_day = "" then 'AND cd_folga = "1" then%>
					<tr bgcolor="<%=str_cor_dia%>" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=str_cor_dia%>');">
						<td align=center><%'=mens%><font color="<%=dia_cor%>"><%=dia_mes%></font></td>
						<td align=Left colspan=3><%=str_msg_ausencia%> <%=nm_dsr%></td>
						<td align=center><%=str_min%></td>
						<td align=center>&nbsp;</td>
						<!--td align=center><img src="imagens/blackdot.gif" width="1" height="20"></td-->
						<td align=left>&nbsp;<%=txt_obs%>&nbsp;<%=msg_2%><%=msg_1%><%=str_desligamento%></td>
					    <%If fechado = "1" Then%>
					    <Td align=center id="no_print"><input type="button" value="   ...   "></td>
					<%Else%>
						<td align=right id="no_print">
						<!--input onclick="javascript:JsInsert('<%=dia_mes%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="I" title="Incluir"> 
						<input onclick="javascript:JsEdit('<%=dia_mes%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="E" title="Editar"-->
						<!--a href="javascript:void(0)" onclick="window.open('empresa/abono/abono_ausencias.asp?cd_funcionario=<%=cd_codigo%>&dt_data=<%=str_dia%>', 'principal', 'width=200, height=400'); return false;"><img src="imagens/ic_clock.gif" border="0"></a-->
						<!--a href="../empresa/abono/abono_ausencias.asp?cd_funcionario=<%=cd_codigo%>&dt_data=<%=str_dia%>" target="_blank" title="Abono"><img src="imagens/ic_clock.gif" border="0"></a-->&nbsp;&nbsp;&nbsp;
						<img src="imagens/ic_del.gif" onclick="javascript:JsDelfolga('<%=cd_codigo%>','<%=cd_cod_folga%>','<%=cd_folga%>','<%=cd_dsr%>')" type="button" value="A" title="Apagar">
						</td>
							<%if tela = "1" Then%>
							<td id="no_print"><img src="imagens/px.gif" width=10 height=0></td>
							<%end if%>
					<%end if%>
	<%'end if
	rs3.movenext
			loop
		'*** limpa dia de folga
		'cd_folga = ""
		dt_data = ""


'***********************************************
'** 3 		/Linhas dos dias em branco\		 ***
'***********************************************

		'***********************************************************
		'*	Verifica se há horas abonadas para os dias trabalhados *
		'***********************************************************
		xsql3 = "up_funcionario_abono @cd_funcionario="&cd_codigo&", @dt_abono='"&mes_sel&"/"&dia_mes&"/"&ano_sel&"'"
		Set rs3 = dbconn.execute(xsql3)
		'response.write(cd_codigo&"-")
		'response.write(dia_mes&"/"&mes_sel&"/"&ano_sel)
		
		Do while not rs3.EOF
			cd_cod_abono = rs3("cd_codigo")
			dt_abono_tempo = rs3("dt_abono_tempo")
			saldo_dia = saldo_dia + int(dt_abono_tempo)
			
			str_abono_tempo = total_mes(dt_abono_tempo)'strsum
			strsun = strsun + int(dt_abono_tempo)
			msg_abono = "&nbsp;&nbsp;&nbsp;&nbsp; A&nbsp;&nbsp;&nbsp; B&nbsp;&nbsp;&nbsp; O&nbsp;&nbsp;&nbsp; N&nbsp;&nbsp;&nbsp; O&nbsp;&nbsp;&nbsp;&nbsp; "
			msg_3 = "+"&total_mes(int(dt_abono_tempo))&"hs"
		rs3.movenext
		loop
	
	
	'*** Procura dia duplicado no banco de dados ***
	'str_duplos_1 = instr(1,duplicados_1, day(dt_data),1)
	'duplicados_1 = duplicados_1&","&day(dt_data)
	if cd_folga = "" Then
		cd_folga = "0"
	End if
	if cd_folga = "" Then
		cd_folga = "0"
	end if
	'***********************************************

	'if data_entrada_day = "" AND cd_folga <> "1" OR dia_mes <> day(dt_data) AND cd_folga <> "1" Then'AND str_duplos_2 = 0 then
if data_entrada_day = "" AND cd_folga < "1" Then%>
					<tr bgcolor="#F5F5F5" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#F5F5F5');">
						<!--td>&nbsp;</td-->
						<td align=center><%'=mens%><font color="<%=dia_cor%>"><%=dia_mes%></font></td>
						<td align=left colspan="3"><%=msg_abono%>&nbsp;</td>
						<!--td align=center><img src="imagens/px.gif" width="1" height="0"></td-->
						<!--td align=center></td-->
						<td align=center><%=str_abono_tempo%><img src="imagens/px.gif" width="1" height="0"></td>
						<td align=center>&nbsp;</td>
						<td align=left>&nbsp;<%=msg_3%><%'=msg_1%><%=str_desligamento%></td>
					<%If fechado = "1" Then%>
					    <Td align=center id="no_print"><input type="button" value="   ...   "></td>
					<%Else%>
						<Td align=left id="no_print">&nbsp;&nbsp;
						<img src="imagens/ic_novo.gif" onclick="javascript:JsInsert('<%=dia_mes%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>','<%=tela%>')" type="button" value="I" title="Incluir"><img src="imagens/px.gif" width="2" height="0"><a href="javascript:void(0)" onclick="window.open('empresa/abono/abono_ausencias.asp?cd_funcionario=<%=cd_codigo%>&dt_data=<%=str_dia%>', 'principal', 'width=200, height=400'); return false;"><img src="imagens/ic_clock.gif" border="0"></a>
						<!--a href="javascript:void(0)" onclick="window.open('empresa/abono/abono_ausencias.asp?cd_funcionario=<%=cd_codigo%>&dt_data=<%=str_dia%>', 'principal', 'width=200, height=400'); return false;"><img src="imagens/ic_clock.gif" border="0"></a-->
						<!--a href="../empresa/abono/abono_ausencias.asp?cd_funcionario=<%=cd_codigo%>&dt_data=<%=str_dia%>" target="_blank" title="Abono"><img src="imagens/ic_clock.gif" border="0"></a-->&nbsp;&nbsp;&nbsp;
						<!--input onclick="javascript:JsEdit('<%=dia_mes%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="E" title="Editar"--> 
						<%if cd_cod_abono <> "" Then%>
						<a href="javascript:void(0)" onclick="window.open('empresa/horas_trabalhadas/horas_trab_cartao_excluir.asp?cd_cod_abono=<%=cd_cod_abono%>&cd_cod_horas=', 'principal', 'width=100, height=300'); return false;"><img src="imagens/ic_del.gif" border="0" Alt="Apagar"><%end if%></a>
						</td>
							<%if tela = "1" Then%>
							<td id="no_print">&nbsp;</td>
							<%end if%>
					<%end if%>	
	<%end if
	%>
					
	</tr>
	</tabela>
					<%
					'strsun = strsun + min
					'strdias = strdias + 1
					'conta_dias = conta_dias + 1
					
				'*** Procura dia duplicado no banco de dados ***
						
						'str_duplos = instr(1,duplicados, data_entrada_day,1)
						'duplicados = duplicados&","&data_entrada_day
				'***********************************************
					
				'*** Limpa algumas variáveis ***
					'duplicados = ""
					'cd_folga = ""
					cd_cod_abono=""
					dt_abono_tempo = ""
					str_abono_tempo = ""
					msg_1 = ""
					msg_2 = ""
					msg_3 = ""
					msg_abono = ""
					cor_saldo_dia = ""
					str_cor_dia = ""
					data_entrada_day = ""
					nm_obs = ""
					strsinal = ""
					strcor = ""
					strst_ = ""
					cd_folga = ""
					str_min = ""
					nm_dsr = ""
					str_desligamento = ""
					dia_mes = dia_mes + 1
					min = "0"

						
						
Loop
							'rs2.close
							'set rs2 = nothing%>
			
			<tr bgcolor=#FFFFFF id="ok_print"><!--espaçador impressão -->
				<td><img src="imagens/px.gif" width=28 height=0></td>
				<td><img src="imagens/px.gif" width=50 height=0></td>
				<td><img src="imagens/px.gif" width=80 height=0></td>
				<td><img src="imagens/px.gif" width=50 height=0></td>
				<td><img src="imagens/px.gif" width=60 height=0></td>
				<td><img src="imagens/px.gif" width=60 height=0></td>
				<td><img src="imagens/px.gif" width=130 height=0></td>
			</tr>
		<tr id="ok_print"><td colspan="100%"><img src="imagens/blakdot.gif" width="100%" height="1"></td></tr>
			</table>	

					
</div></td></tr>

</td></tr>				
				<tr bgcolor=#FFFFFF id="no_print"><!--Separador e limitador de espaços-->
						<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<td colspan="3" align="center">&nbsp;Total trabalhado no mês: </td>
					<td align="center"><b><%=total_mes(strsun)%></b></td>
					<td align="center"><img src="imagens/px.gif" width=1 height=0></td>
					<td id="no_print"><img src="imagens/px.gif" width=1 height=0></td>
					<td><img src="imagens/px.gif" width=1 height=0></td>
				</tr>
				<!--tr bgcolor="<%=strcor%>"-->
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<!--td><img src="imagens/px.gif" width=1 height=0></td>
					<td colspan="3" align="center">&nbsp;Saldo total no mês: </td>
					<td align="center"><b></b></td>
					<td align="center"><%'=total_mes(saldo_mes)%><%=total_mes(saldo_dia)%></td>
					<td id="no_print"><img src="imagens/px.gif" width=1 height=0></td>
					<td><img src="imagens/px.gif" width=1 height=0></td>
				</tr-->
				<tr><td colspan="100%">&nbsp;</td></tr>
			<tr id="ok_print"><td colspan="100%"><img src="imagens/blakdot.gif" width="100%" height="1"></td></tr>
					<!--tr bgcolor=#FFFFFF id="no_print"><Separador e limitador de espaços>
						<td><img src="imagens/blackdot.gif" width=30 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=110 height=0></td>
						<td><img src="imagens/blackdot.gif" width=110 height=0></td>
						<td><img src="imagens/blackdot.gif" width=100 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=150 height=0></td>
						<td><img src="imagens/blackdot.gif" width=60 height=0></td>
						<td><img src="imagens/blackdot.gif" width=13 height=0></td>
					</tr-->
				<tr bgcolor=#FFFFFF id="no_print"><!--Separador e limitador de espaços-->
					<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
					
				<%If fechado <> "1" Then%>
				<tr id="no_print">
					<td align=center>Dia</td>
					<td align=center>Entrada</td>
					<!--td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td align=center >Saída Intervalo</td>
					<!--td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td align=center>Entrada Intervalo</td>
					<td align=center>Saida</td>
					<!--td align=center>Folga/DSR/Falta</td-->
					<td align=center>Folga/Falta</td>
					<td align="center">Observações</td>
					<td align=center><img src="imagens/blackdot.gif" width=1 height=0></td>					
				</tr>
				<tr id="no_print">
			   <%end if
			   if acao = "editar" OR acao = "copiar" then
					xsql3 = "up_Hora_Trabalhada_lista_2 @cd_funcionario="&cd_codigo&",@dia="&dia_sel&",@mes="&mes_sel&",@ano='"&ano_sel&"'"
					
						Set rs3 = dbconn.execute(xsql3)
						if NOT rs3.EOF Then
						cd_cod_horas = rs3("cd_codigo")
						nm_obs = rs3("nm_obs")
				'Dia Entrada
					dia_entrada = day(rs3("data_entrada"))
					mes_entrada = month(rs3("data_entrada"))
					ano_entrada = year(rs3("data_entrada"))
					'Entrada
						hora_entrada = hour(rs3("data_entrada"))
						minuto_entrada = minute(rs3("data_entrada"))
							entrada = hora_entrada + minuto_entrada
					'Saída intervalo
						hora_sai_intr = hour(rs3("data_saida_interv"))
						minuto_sai_intr = minute(rs3("data_saida_interv"))
							saida_intervalo = hora_sai_intr + minuto_sai_intr
					'Entrada intervalo
						hora_entr_intr = hour(rs3("data_entr_interv"))
						minuto_entr_intr = minute(rs3("data_entr_interv"))
							entrada_intervalo = hora_entr_intr + minuto_entr_intr
				'Dia Saida						
					dia_saida = day(rs3("data_saida"))
					mes_saida = month(rs3("data_saida"))
					ano_saida = year(rs3("data_saida"))
					'Saída
						hora_saida = hour(rs3("data_saida"))
						minuto_saida = minute(rs3("data_saida"))
							saida = hora_saida + minuto_saida
						
						
						dt_entrada = dia_entrada&"/"&mes_entrada&"/"&ano_entrada&" "&hora_entrada&":"&minuto_entrada
						dt_saida = dia_saida&"/"&mes_saida&"/"&ano_saida&" "&hora_saida&":"&minuto_saida
					
						minutos_trab = DateDiff("n",cdate(dt_entrada),cdate(dt_saida))
						
						Else
						data_entrada = "01-01-5000"
						data_saida = "01-01-5000"
						End if
				Else
					mes_entrada = mes_sel
					ano_entrada = ano_sel
					
					mes_saida = mes_sel
					ano_saida = ano_sel
					acao = "inserir"
				End if
				
				
				cd_carga_horaria = cd_carga_horaria * 60
				cd_carga_diaria = cd_carga_diaria * 60
				
					'Verifica se há carga horaria para o mes de trabalho
			'			xsql = "SELECT * From TBL_carga_horaria where cd_funcionario = '"&cd_codigo&"' AND cd_mes = '"&strmes_sel&"' AND cd_ano = '"&strano_sel&"'"
			'			Set rs = dbconn.execute(xsql)
			'			Do while not rs.EOF
			'			
			'			'cd_carga_horaria = rs("cd_codigo")
			'			cd_carga_horaria = rs("cd_carga_horaria")
			'			'response.write("teste")
			'			'cd_carga_horaria = replace(cd_carga_horaria,":",",")
			'				'if cd_carga_horaria <> "" Then
			'					cd_carga_horaria = cd_carga_horaria '* 60
			'				'end if
			'			
			'			rs.movenext
			'			Loop
				
				excedente = strsun - cd_carga_horaria
				
				if excedente = "" Then
					excedente = 0
				end if
				If fechado <> "1" Then
				
				'******************************************************
				'*** 		Dados do campo dia_entrada	 			***
				'******************************************************
				
					if dia_entrada="" Then
					str_entrada = dia_sel
					elseif dia_entrada > 0 AND dia_entrada < 10 Then
					'response.write("0"&dia_entrada)
					else
					str_entrada = dia_entrada
					end if
					
						'*** Em caso de cópia, 
						if acao = "copiar" then
						str_entrada = ""
						end if
				'******************************************************
				%>
				<form name="form" action="empresa/horas_trabalhadas/horas_trab_cartao_acao.asp" id="form" method="post">
				<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
				<input type="hidden" name="cd_status" value="<%=cd_status%>">
				<input type="hidden" name="cd_unidades" value="<%=cd_unidades%>">
				<input type="hidden" name="cd_funcao" value="<%=cd_funcao%>">
				<input type="hidden" name="mesentrada" value="<%=mes_sel%>">
				<input type="hidden" name="anoentrada" value="<%=ano_sel%>">
				<input type="hidden" name="cd_cod_horas" value="<%=cd_cod_horas%>">
				<input type="hidden" name="acao" value="<%=acao%>">
				<input type="hidden" name="cd_excedente" value="<%=excedente%>">
				<input type="hidden" name="cd_minutos_trab" value="<%=minutos_trab%>">
				<input type="hidden" name="total_geral_trabalhado" value="<%=strsun%>">
				<input type="hidden" name="tela" value="<%=tela%>">
								 
					<td align=center><input type="txt" name="diaentrada" size=2 maxlength="2" value="<%=str_entrada%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center ><input type="txt" name="horaentrada_hr" size=2 maxlength="2" value="<%if hora_entrada <> "" and  hora_entrada < 10 then%><%response.write("0"&hora_entrada)%><%Else%><%=hora_entrada%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="txt" name="horaentrada_min" size=2 maxlength="2" value="<%if minuto_entrada <> "" and  minuto_entrada < 10 then%><%response.write("0"&minuto_entrada)%><%Else%><%=minuto_entrada%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<!--td align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
				  	<td align=center>das <input type="txt" name="intsaida_hr" size=2 maxlength="2" value="<%if saida_intervalo = entrada then%><%elseif hora_sai_intr <> "" and  hora_sai_intr < 10 then%><%response.write("0"&hora_sai_intr)%><%Else%><%=hora_sai_intr%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:
										 <input type="txt" name="intsaida_min" size=2 maxlength="2" value="<%if saida_intervalo = entrada then%><%elseif minuto_sai_intr <> "" and  minuto_sai_intr < 10 then%><%response.write("0"&minuto_sai_intr)%><%Else%><%=minuto_sai_intr%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<!--td align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
				  	<td align=center>às <input type="txt" name="intentrada_hr"  size=2 maxlength="2" onKeyPress="ChecarTAB();"  value="<%if entrada_intervalo = entrada then%><%elseif hora_entr_intr <> "" and  hora_entr_intr < 10 then%><%response.write("0"&hora_entr_intr)%><%Else%><%=hora_entr_intr%><%end if%>" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:
										<input type="txt" name="intentrada_min" size=2 maxlength="2" value="<%if entrada_intervalo = entrada then%><%elseif minuto_entr_intr <> "" and  minuto_entr_intr < 10 then%><%response.write("0"&minuto_entr_intr)%><%Else%><%=minuto_entr_intr%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center><input type="txt" name="horasaida_hr" size=2 maxlength="2" value="<%if hora_saida <> "" and  hora_saida < 10 then%><%response.write("0"&hora_saida)%><%Else%><%=hora_saida%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="txt" name="horasaida_min" size=2 maxlength="2" value="<%if minuto_saida <> "" and  minuto_saida < 10 then%><%response.write("0"&minuto_saida)%><%Else%><%=minuto_saida%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center><input type="radio" name="cd_folga" value="1" size=1 maxlength="1"><!--***Folga*** -->
									 <!--input type="radio" name="cd_folga" value="2"  size=1 maxlength="1"--><!--***Folga DSR*** -->
									 <input type="radio" name="cd_folga" value="3"  size=1 maxlength="1"><!--***Falta*** -->
								</td>
					<td align=center><input type="text" name="nm_obs" size="20" value="<%=nm_obs%>"></td>
					<td align=center id="no_print"><input type="submit" value="<%=nm_acao%>">
				</tr>
				<%else%>
				<tr id="no_print">
					<td colspan="100%">&nbsp;<br><br></td>
				</tr>
				<%end if%>
				<tr bgcolor=#FFFFFF id="no_print"><!--Separador e limitador de espaços-->
						<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
				<tr bgcolor="<%=strcor%>" id="no_print">
					<td colspan="100%">&nbsp;</td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td><img src="imagens/px.gif" width=1 height=0></td><br>
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td colspan="2">&nbsp;Total trabalhado: </td>
					<td align="right"><b><%=total_mes(strsun)%> </b></td>
					<td align="left"><b>horas</b></td>
					<td colspan="2"><img src="imagens/px.gif" width=1 height=0><!--input type="submit" value="<%'=acao%>"--></td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td colspan="2">&nbsp;Carga Horária: </td>
					
					<td align="right" id="ok_print"><b><%=total_mes(cd_carga_horaria)%> </b></td>
					<td align="left" id="ok_print"><b>hs/mês</b></td>
					
					<td align="right" id="no_print"><%If fechado <> "1" Then%><input type="text" name="cd_carga_horaria" size="6" maxlength="6" value="<%end if%><%=total_mes(cd_carga_horaria)%><%If fechado <> "1" Then%>"><%end if%></td>
					<td align="left" id="no_print"><b>hs/mês</b></td>
					<td colspan="3"><img src="imagens/px.gif" width=1 height=0><!--input type="submit" value="<%'=acao%>"--></td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td colspan="2">&nbsp;Carga diária: </td>
					<td align="right" id="ok_print"><b><%=total_mes(cd_carga_diaria)%></b></td>
					<td align="left" id="ok_print"><b>hs/dia</b></td>
					
					<td align="right" id="no_print"><%If fechado <> "1" Then%><input type="text" name="cd_carga_diaria" size="6" maxlength="6" value="<%end if%><%=total_mes(cd_carga_diaria)%><%If fechado <> "1" Then%>"><%end if%></b></td>
					<td align="left" id="no_print"><b>hs/dia</b></td>
					<td colspan="3"><img src="imagens/px.gif" width=1 height=0><!--input type="submit" value="<%'=acao%>"--></td>
				</tr></form>
				
				<tr bgcolor="<%=strcor%>">
					<form action="../intranet/empresa/banco_horas/banco_horas_acao_teste.asp" name="banco_horas" id="banco_horas">
					<td><img src="imagens/px.gif" width=1 height=0></td><input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
					<td><img src="imagens/px.gif" width=1 height=0></td><input type="hidden" name="anoentrada" value="<%=ano_sel%>">
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td>--><input type="hidden" name="mesentrada" value="<%=mes_sel%>">
					<td colspan="2">&nbsp;Saldo do mês: </td><input type="hidden" name="cd_saldo_mes" value="<%=excedente%>">
					<td align="right"><b><%=total_mes(excedente)%> </b></td><input type="hidden" name="manual" value="1">
					<td align="left"><b>h/mês</b></td>
					<td colspan="3"><img src="imagens/px.gif" width=1 height=0><%If fechado <> "1" Then%><!--input type="submit" name="banco_horas" value="Banco de horas" id="no_print"--><%end if%></td>
				</tr></form>
				<tr bgcolor=#FFFFFF><!--Separador e limitador de espaços-->
					<td colspan="100%">
						<p>&nbsp;</p>
						<p>&nbsp;</p>
					</td>
				</tr>
				<tr bgcolor=#FFFFFF><!--Separador e limitador de espaços-->
					<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
				<%xsql4 = "SELECT SUM(cd_horas_excedentes) AS horas_extras FROM TBL_banco_horas WHERE (cd_funcionario = '"&cd_codigo&"') AND (dt_mes <= '"&mes_sel&"') AND (dt_ano <= '"&ano_sel&"')"
				Set rs4 = dbconn.execute(xsql4)
				do while not rs4.EOF
				horas_extras = rs4("horas_extras")
				rs4.movenext
				loop
				%>
				<tr>
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<td><img src="imagens/px.gif" width=1 height=0></td>
					<td colspan="2" align="left">&nbsp;Saldo total atual<br>&nbsp;do Banco de horas:</td>
					<td align="right" valign="bottom"><%=total_mes(horas_extras)%></td>
					<td align="left" valign="bottom"><b> horas</b></td>
				</tr>
				
				<tr bgcolor=#FFFFFF><!--Separador e limitador de espaços-->
					<td colspan="100%"><img src="imagens/blackdot.gif" width=100% height=0></td>
				</tr>
				<tr id="no_print">
					<td colspan = "100%"><br><br><br>
					Excedente: <b><%=total_mes(excedente)%>hs </b>  (<%=excedente%> min)<br>
					Total trabalhado: <b><%=(total_mes(strsun))%> hs </b> - (<%=strsun%> min)<br>
					Total de dias trabalhados: <b><%=conta_dias%><br></b>
					Quantidade de dias no mes: <b><%=ultimo_dia%></b><br>
					</td>
				</tr>
				<tr><td>&nbsp;<br><br><br></td></tr>
					<%End if%>
			</table>
		</td>
	</tr>
</table>


<SCRIPT LANGUAGE="javascript">

function JsInsert(cod_a1, cod_a2, cod_a3, cod_a4, cod_a5)
{
		document.location.href('empresa_funcionarios_horas.asp?dia_sel='+cod_a1+'&mes_sel='+cod_a2+'&ano_sel='+cod_a3+'&cd_codigo='+cod_a4+'&tela='+cod_a5+'&acao=inserir');
}
function JsEdit(cod_a, cod_b, cod_c, cod_d, cod_e)
{
		document.location.href('empresa_funcionarios_horas.asp?dia_sel='+cod_a+'&mes_sel='+cod_b+'&ano_sel='+cod_c+'&cd_codigo='+cod_d+'&tela='+cod_e+'&acao=editar');
}
function JsDelete(cod1, cod2, cod3, cod4, cod5, cod6)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('empresa/horas_trabalhadas/horas_trab_cartao_acao.asp?cd_codigo='+cod1+'&cd_cod_horas='+cod2+'&acao=excluir&mesentrada=<%=mes_entrada%>&anoentrada=<%=ano_entrada%>&menos_min='+cod3+'&cd_excedente=<%=excedente%>&cd_minutos_trab='+cod4+'&tela='+cod5+'&cd_cod_abono='+cod6+'');
	  }
}
function JsCopy (cod_a, cod_b, cod_c, cod_d, cod_e)
{
		document.location.href('empresa_funcionarios_horas.asp?dia_sel='+cod_a+'&mes_sel='+cod_b+'&ano_sel='+cod_c+'&cd_codigo='+cod_d+'&tela='+cod_e+'&acao=copiar');
}
function JsDelfolga(cod1, cod2, cod3, cod4)
{
  if (confirm ("Tem certeza que deseja excluir a folga?"))
	  {
		document.location.href('empresa/horas_trabalhadas/horas_trab_cartao_acao.asp?cd_codigo='+cod1+'&cd_cod_folga='+cod2+'&cd_folga='+cod3+'&acao=excluir&mesentrada=<%=mes_entrada%>&anoentrada=<%=ano_entrada%>');
	  }
}
</SCRIPT>
<%
					min = ""
					strsun = 0
					str_duplos = ""
					duplicados = ""
					negativo_dia = ""
					cd_carga_horaria = ""%>
					<%=mes_sel%>/31/<%=ano_sel%>
</body>
</html>
