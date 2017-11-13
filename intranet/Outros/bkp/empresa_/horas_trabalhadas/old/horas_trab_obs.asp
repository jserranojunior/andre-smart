<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--include file="../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<style type="text/css">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #000000; }  
.print {color: #000000;font-family: verdana;font-size:9px;text-decoration:none;}
-->
</style>

<%
mes_prod = month(now)
ano_prod = year(now)

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")
data_atual = ano_sel&mes_sel

cd_codigo = request("cd_codigo")
acao = request("acao")

If acao = "excluir" Then
xsql ="up_cartao_horas_excluir @cd_codigo='"&strhora&"',@cd_horas="
dbconn.execute xsql
response.write("OK, Excluído")
'response.redirect "Funcionarios_cartao.asp?cod="&strcod&"&mes_sel="&strmes_sel&""

end if


Function formatahora(min)

Hora = Int(min/60)   '60 minutos
Minutos = min Mod 60 '60 segundos
NovaHora = Right("0" & Hora,2) & ":" & Right("0" & Minutos,2)
formatahora = NovaHora

End Function


Function total_mes(min)

Hora = Int(min/60)   '60 minutos
Minutos = min Mod 60 '60 segundos
NovaHora = Right("" & Hora,4) & ":" & Right("0" & Minutos,2)
total_mes = NovaHora
End Function



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

xsql_verif = "up_Verifica_fechamento_mes @cd_mes = '"&strmes_sel&"', @cd_ano = '"&strano_sel&"'"
Set rs_verif = dbconn.execute(xsql_verif) 'Verifica o fechamento do mes

Do While Not rs_verif.EOF

strclose = rs_verif("nm_fechado")

If strclose <> "" Then
encerra = "encerrado"
fechado = "1"

End If
rs_verif.movenext

loop
%>
<html>
<head>
	<TITLE> New Document </TITLE>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
</head>
<LINK href="css/estilo.css " type=text/css rel=stylesheet >

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
<table width="500" cellspacing="0" cellpadding="0" border="1" class="textoPadrao" align="center">
	<tr>
		<td align="center">
			<table width="100%" cellspacing="1" cellpadding="0" border="1" class="textoPadrao">
					<tr bgcolor=#cococo><!--Separador e limitador de espaços-->
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=150 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=75 height=0></td>
					</tr>
					<tr>
						<td colspan="100%"><b>FUNCIONÁRIOS - <font color="red">Banco de horas</font></b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
				<%'*********************************************************
				  '*				      1ª Parte	 					  '*
				  '*********************************************************
					if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then%>
					
					<tr>
					    <td colspan="10" align="center">Selecione um período abaixo:</td>
					  
						
					</tr>
					<tr align="center">
					    <td colspan="10"><%
						if mes_prod = "1" Then
						mes_ant = 12
						ano_ant = ano_prod - 1
						Else 
						mes_ant = mes_prod - 1
						ano_ant = ano_prod
						end if
						%>
						<a href=empresa_funcionarios_obs.asp?mes_sel=<%=mes_ant-1%>&ano_sel=<%=ano_ant%>><%=mes_selecionado(mes_ant-1)%>/<%=ano_ant%></a></td>
					</tr>
					<tr align="center">
					    <td colspan="10"><%
						if mes_prod = "1" Then
						mes_ant = 12
						ano_ant = ano_prod - 1
						Else 
						mes_ant = mes_prod - 1
						ano_ant = ano_prod
						end if
						%>
						<a href=empresa_funcionarios_obs.asp?mes_sel=<%=mes_ant%>&ano_sel=<%=ano_ant%>><%=mes_selecionado(mes_ant)%>/<%=ano_ant%></a></td>
					</tr>
					<tr align="center">
					    <td colspan="10"><a href=empresa_funcionarios_obs.asp?mes_sel=<%=mes_prod%>&ano_sel=<%=ano_prod%>><b><%=mes_selecionado(mes_prod)%>/<%=ano_prod%></b></a></td>
					</tr>
					<tr align="center">
					    <td colspan="10"><%
						mes_prox = month(now)+ 1
						ano_prox = year(now)
						
						if mes_prox = "13" Then
						mes_prox = 1
						ano_prox = ano_prox + 1
						end if
						%>
						<a href=empresa_funcionarios_horas?mes_sel=<%=mes_prox%>&ano_sel=<%=ano_prox%>><%=mes_selecionado(mes_prox)%>/<%=ano_prox%></a>
						</td>
					</tr>
					<tr>
					    <td colspan="10">&nbsp;</td>
					</tr>
				<%'*********************************************************
				  '*				      2ª Parte	 					  '*
				  '*********************************************************
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
			<table width="100%" cellspacing="1" cellpadding="1" border="0" class="textoPadrao">
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="4">[<a href="empresa_funcionarios_obs.asp">Mudar Mês</a>] &nbsp;&nbsp; <b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td colspan="4"><br>&nbsp;</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td><b>Código</b></td>
						<td><b>Matrícula</b></td>
						<td><b>Nome</b></td>
					</tr>
					
					<%conta_coop = "1"
					condicao = "Where cd_regime_trabalho=1"
					xsql = "up_funcionarios_lista @cd_codigo='', @ordem='nm_nome,nm_sobrenome', @condicao='"&condicao&"'"
					Set rs = dbconn.execute(xsql)
					
					Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					cd_unidades = rs("cd_unidades")
					dt_demissao = rs("dt_demissao")
					cd_status = rs("cd_status")

					if dt_demissao <= data_atual Then
					cor_font = "#afafaf"
					end if
					%>
					<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa_funcionarios_obs.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&cd_codigo=<%=cd_codigo%>');" onmouseout="mOut(this,'');">
						<td>&nbsp;</td>
						<td><font color="<%=cor_font%>"><%=cd_codigo%></font></td>
						<td><font color="<%=cor_font%>"><%=cd_matricula%></font></td>
						<td><font color="<%=cor_font%>"><%=nm_nome%>&nbsp;<%=nm_sobrenome%><%'=cd_status%></font></td>
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
				<%'*********************************************************
				  '*				      3ª Parte	 					  '*
				  '*********************************************************
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo <> "" Then%>
					
					<%xsql = "up_funcionarios_lista @cd_codigo='"&cd_codigo&"', @ordem='nm_nome, nm_sobrenome', @condicao=''"
					Set rs = dbconn.execute(xsql)
					
					'Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					nm_foto = rs("nm_foto")
					cd_unidades = rs("cd_unidades")
					dt_demissao = rs("dt_demissao")
					cd_status = rs("cd_status")
					cd_funcao = rs("cd_funcao")
					cd_carga_horaria = rs("cd_carga_horaria")
					
					
					if dt_demissao <= data_atual Then
					status = "(desligado)"
					end if
					%>
					
					<tr >
						<td  align=center colspan="2"><%=nm_nome%>&nbsp;<%=nm_sobrenome%><br><%'=cd_status%></td>
			    		<td align=center>&nbsp;</td>
						<td align=center colspan=4>
							<% If nm_foto <> "" then%>
							<img src="foto/funcionarios/<%=nm_foto%>" width=100 >
							<%Else%>
							<img src="imagens/Vdlap2p.gif" width=100>
							<br><br>
							<%End if%>
						
			  		</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td colspan="9">[<a href="empresa_funcionarios_obs.asp">Mudar Mês</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp;[<a href="empresa_funcionarios_obs.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>">Mudar Cooperado]</td>
					</tr>
					
					<tr>
						<td colspan="100%"><br>&nbsp;&nbsp;<b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b></td>
					</tr>
					
					
					<!--tr bgcolor=#cococo--><!--Separador e limitador de espaços-->
						<!--td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=150 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=75 height=0></td>
					</tr-->
				</table>
				<table cellspacing="1" cellpadding="1" border="1" class="textoPadrao">
				<tr bgcolor="#cococo">
					<td align=center>Dia</td>
					<td align=center>Entrada</td>
					<!--td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td align=center>Intervalo</td>
					<!--td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td align=center>Entrada Intervalo</td-->
					<td align=center>Saida</td>
					<td align=center>Total</td>
					<td align=center>Banco de horas</td>
					<td align="center">Observações</td>
					<td align=center>Ações</td>
				<tr>
					<%
					primeiro_dia = 1
					dia_mes = primeiro_dia
					ultimo_dia = UltimoDiaMes(mes_sel,ano_sel)
					'ultimo_dia = 15
					'************************************************************
								'Verifica se há carga horaria para o mes de trabalho
							xsql = "SELECT * From TBL_carga_horaria where cd_funcionario = '"&cd_codigo&"' AND cd_mes = '"&strmes_sel&"' AND cd_ano = '"&strano_sel&"'"
							Set rs = dbconn.execute(xsql)
							Do while not rs.EOF
							
							'cd_carga_horaria = rs("cd_codigo")
							cd_carga_horaria = rs("cd_carga_horaria")
							'response.write("teste")
							'cd_carga_horaria = replace(cd_carga_horaria,":",",")
								'if cd_carga_horaria <> "" Then
									cd_carga_horaria = cd_carga_horaria '* 60
									cd_carga_horaria = cd_carga_horaria / 60
								'end if
							
							rs.movenext
							Loop
						'************************************************************
						
do while not dia_mes = ultimo_dia + 1
						
					xsql2 = "up_Hora_Trabakhada_por_cd_funcionario_4 @cd_funcionario="&cd_codigo&",@dia="&dia_mes&",@mes="&mes_sel&",@ano='"&ano_sel&"'"
					Set rs2 = dbconn.execute(xsql2)
					
					do while not rs2.eof
					
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
									'mens = "(a) "
								elseif hour(data_sai)< hour(data_entrada) Then ' **** intervalo após a meia noite ***
									min1 = datediff("n",cdate(data_entrada),cdate(data_entrada_month&"/"&data_entrada_day&"/"&data_entrada_year&" "&"23:59"))
									min2 = datediff("n",cdate(data_entrada_month&"/"&data_entrada_day +1 &"/"&data_entrada_year&" 0:0"),cdate(data_sai))
									min3 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = 1 + min1 + min2 + min3
									'mens = "(1) "
								Elseif hour(data_entr)< hour(data_sai) Then ' ***** Intervalo em dias diferentes
									min0 = datediff("n",cdate(data_entrada),cdate(data_sai))
									min1 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = min0 + min1
									'mens = "(2) "
								Elseif hour(data_saida)< hour(data_entrada) Then ' ***** Saída no outro dia
									min0 = datediff("n",cdate(data_entrada),cdate(data_sai))
									min1 = datediff("n",cdate(data_entr),cdate(data_entr_month&"/"&data_entr_day&"/"&data_entr_year&" 23:59"))
									min2 = datediff("n",cdate(data_saida_month&"/"&data_saida_day&"/"&data_saida_year&" 0:0"),cdate(data_saida))
									min = 1 + min0 + min1 + min2
									'mens = "(3) "
								Else
									min1 = datediff("n",cdate(data_entrada),cdate(data_sai))
									min2 = datediff("n",cdate(data_entr),cdate(data_saida))
									min = min1 + min2
									'mens = "(0) "
								end if
								
								'******** Tempo de intervalo - Não desconta para quem faz 12x36  *******
								tempo_interv = datediff("n",cdate(data_sai),cdate(data_entr))
								If tempo_interv = 60 AND cd_carga_horaria = 156 then
								intervalo_livre = "60"
								min = min + intervalo_livre
								End if
						'**********************************************************************************
						
						'*** Procura jornada de trabalho negativa e avisa ***
						negativo = Instr(1,min,"-",1)'Compara se o caracter "-" aparece na string min
						If negativo = "1" Then
						strst_ = "0" 
						strsinal = "ERRO!"
						end if
						'****************************************************************************
						
						
						IF strst_ = "" Then
							strcor = "#eaeaea"
						ElseIf strst_ ="5"Then
							strcor = "#E0FFFF"
						Elseif strst_ = "0" Then
							strcor = "#ff8080"
						End If 
						
						if str_duplos <> 0 then
						str_cor_dia = "#FF0000"
						end if
					
						'*** Subtotal das horas ***
						strsun = strsun + min 
						'**************************
					'*** Linhas dos dias do mes trabalhado ***%>
					<tr bgcolor="<%=strcor%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=strcor%>');">
						<td align=center width=172><%=mens%><font color=<%=str_cor_dia%>><%=data_entrada_day%></font></td>
						<td align=center width=172><font color=<%=str_cor_dia%>><%=data_entrada_hour%>:<%=data_entrada_minute%>:<%=data_entrada_second%></font></td>
						<td align=center width=172><font color=<%=str_cor_dia%>><%if data_sai = data_entrada Then%><%response.write("&nbsp;-")%><%else%><%=data_sai_hour%>:<%=data_sai_minute%> às <%=data_entr_Hour%>:<%=data_entr_minute%><%end if%></font></td>
						<td align=center width=172><font color=<%=str_cor_dia%>><%=data_saida_hour%>:<%=data_saida_minute%>:<%=data_saida_second%></font></td>
						<td width=200 align=center><font color=<%=str_cor_dia%>><%if strsinal = "" Then%><%=total_mes(min)%><%else%><%=strsinal%><%end if%></font></td>
						<td align=center width=172><font color=<%=str_cor_dia%>>&nbsp;<%=BH%></font></td>
						<td align=left width=172>&nbsp;<font color=<%=str_cor_dia%>><%=nm_obs%></font></td>
					<%If fechado = "1" Then%>
					    <Td width=200 align=center><input type="button" value="   ...   "></td>
					<%Else
					cd_carga_horaria = cd_carga_horaria'*60
					excedente = strsun - cd_carga_horaria%>
					
						<td width=200 align=right>
						<!--input onclick="javascript:JsInsert('<%=data_entrada_day%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="I" title="Incluir"-->						
						<input onclick="javascript:JsEdit('<%=data_entrada_day%>','<%=data_entrada_month%>','<%=data_entrada_year%>','<%=cd_codigo%>')" type="button" value="E" title="Editar"> 
						<input onclick="javascript:JsDelete('<%=cd_codigo%>','<%=cd_cod_horas%>','<%=min%>','<%=excedente%>')" type="button" value="A" title="Apagar"> 
						</td>
					<%end if%>
					
					<%rs2.movenext
					loop%>
					
	<%'*** Linhas dos dias em branco ***
	if data_entrada_day = "" then%>
					<tr bgcolor="#F5F5F5" onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#F5F5F5');">
						<td align=center width=172><%=mens%><%=dia_mes%></td>
						<td align=center width=172>&nbsp;</td>
						<td align=center width=5><img src="../imagens/px.gif" width="1" height="0"></td>
						<td align=center width=172>&nbsp;</td>
						<td align=center width=5><img src="../imagens/px.gif" width="1" height="20"></td>
						<!--td align=center width=172>&nbsp;</td>
						<td align=center width=172>&nbsp;</td>
						<td align=center width=5><img src="../imagens/px.gif" width="1" height="20"></td-->
						<td width=200 align=center>&nbsp;</td>
						<td width=200 align=left>&nbsp;</td>
					    <Td width=200 align=left>
						<input onclick="javascript:JsInsert('<%=dia_mes%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="I" title="Incluir"> 
						<input onclick="javascript:JsEdit('<%=dia_mes%>','<%=mes_sel%>','<%=ano_sel%>','<%=cd_codigo%>')" type="button" value="E" title="Editar"> 
						</td>
						
	<%end if%>
					
	</tr>
					<%
					'strsun = strsun + min 
					'strdias = strdias + 1
					'conta_dias = conta_dias + 1
					
				'*** Procura dia duplicado no banco de dados ***
						
						'str_duplos = instr(1,duplicados, data_entrada_day,1)
						'duplicados = duplicados&","&data_entrada_day
				'***********************************************
					
				'*** Limpa algumas variáveis ***
					str_cor_dia = ""
					data_entrada_day = ""
					nm_obs = ""
					strsinal = ""
					strcor = ""
					strst_ = ""

					dia_mes = dia_mes + 1

						
						
Loop
							rs2.close
							set rs2 = nothing
					%>
			
				<tr bgcolor="<%=strcor%>">
					<!--td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td colspan="3" align="center">&nbsp;Total trabalhado no mês: </td>
					<td align="center"><b><%=total_mes(strsun)%></b></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
				</tr>
					<tr bgcolor=#FFFFFF><!--Separador e limitador de espaços-->
						<td><img src="imagens/blackdot.gif" width=30 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=110 height=0></td>
						<td><img src="imagens/blackdot.gif" width=110 height=0></td>
						<td><img src="imagens/blackdot.gif" width=100 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=150 height=0></td>
						<td><img src="imagens/blackdot.gif" width=20 height=0></td>
					</tr>
				<%If fechado <> "1" Then%>
				<tr>
					<td align=center>Dia</td>
					<td align=center>Entrada</td>
					<!--td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td align=center >Saída Intervalo</td>
					<!--td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
					<td align=center>Entrada Intervalo</td>
					<td align=center>Saida</td>
					<td align=center>.</td>
					<td align="center">Observações</td>
					<td width=5 align=center><img src="imagens/blackdot.gif" width=1 height=0></td>					
				<tr>
			   <%if acao = "editar" then
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
				
				if excedente="" Then
					excedente = 0
				end if
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
								 
					<td align=center class="txt_cinza"><input type="txt" name="diaentrada" size=2 maxlength="2" value="<%if dia_entrada="" Then%><%=dia_sel%><%elseif dia_entrada > 0 AND dia_entrada < 10 Then%><%response.write("0"&dia_entrada)%><%else%><%=dia_entrada%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center ><input type="txt" name="horaentrada_hr" size=2 maxlength="2" value="<%if hora_entrada <> "" and  hora_entrada < 10 then%><%response.write("0"&hora_entrada)%><%Else%><%=hora_entrada%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="txt" name="horaentrada_min" size=2 maxlength="2" value="<%if minuto_entrada <> "" and  minuto_entrada < 10 then%><%response.write("0"&minuto_entrada)%><%Else%><%=minuto_entrada%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<!--td align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
				  	<td align=center>das <input type="txt" name="intsaida_hr" size=2 maxlength="2" value="<%if saida_intervalo = entrada then%><%elseif hora_sai_intr <> "" and  hora_sai_intr < 10 then%><%response.write("0"&hora_sai_intr)%><%Else%><%=hora_sai_intr%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:
										 <input type="txt" name="intsaida_min" size=2 maxlength="2" value="<%if saida_intervalo = entrada then%><%elseif minuto_sai_intr <> "" and  minuto_sai_intr < 10 then%><%response.write("0"&minuto_sai_intr)%><%Else%><%=minuto_sai_intr%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<!--td align=center><img src="imagens/blackdot.gif" width=1 height=0></td-->
				  	<td align=center>às <input type="txt" name="intentrada_hr"  size=2 maxlength="2" onKeyPress="ChecarTAB();"  value="<%if entrada_intervalo = entrada then%><%elseif hora_entr_intr <> "" and  hora_entr_intr < 10 then%><%response.write("0"&hora_entr_intr)%><%Else%><%=hora_entr_intr%><%end if%>" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:
										<input type="txt" name="intentrada_min" size=2 maxlength="2" value="<%if entrada_intervalo = entrada then%><%elseif minuto_entr_intr <> "" and  minuto_entr_intr < 10 then%><%response.write("0"&minuto_entr_intr)%><%Else%><%=minuto_entr_intr%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center><input type="txt" name="horasaida_hr" size=2 maxlength="2" value="<%if hora_saida <> "" and  hora_saida < 10 then%><%response.write("0"&hora_saida)%><%Else%><%=hora_saida%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="txt" name="horasaida_min" size=2 maxlength="2" value="<%if minuto_saida <> "" and  minuto_saida < 10 then%><%response.write("0"&minuto_saida)%><%Else%><%=minuto_saida%><%end if%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center><!--input type="checkbox" name="cd_folga" size=1 maxlength="1" value="1"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 1)" onFocus="PararTAB(this);"></td-->
					<td align=center><input type="text" name="nm_obs" size="20" value="<%=nm_obs%>"></td>
					<td align=center><input type="submit" value="<%=acao%>">
				</tr>
				<%else%>
				<tr>
					<td colspan="100%">&nbsp;<br><br></td>
				</tr>
				<%end if%>
				<tr bgcolor="<%=strcor%>">
					<td colspan="100%">&nbsp;</td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td><img src="imagens/blackdot.gif" width=1 height=0></td><br>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td>&nbsp;Total trabalhado no mês: </td>
					<td align="right"><b><%=total_mes(strsun)%> h/mês</b></td>
					<td colspan="3"><img src="imagens/blackdot.gif" width=1 height=0><!--input type="submit" value="<%'=acao%>"--></td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td>&nbsp;Carga Horária: </td>
					<td align="right"><!--b><%'=total_mes(cd_carga_horaria)%> h/mês</b-->
					<input type="text" name="cd_carga_horaria" size="6" maxlength="6" value="<%=total_mes(cd_carga_horaria)%>"> <b>h/mês</b></td>
					<td colspan="3"><img src="imagens/blackdot.gif" width=1 height=0><!--input type="submit" value="<%'=acao%>"--></td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td><img src="imagens/blackdot.gif" width=1 height=0></td>
					<td>&nbsp;Saldo do mês: </td>
					<td align="right"><b><%=total_mes(excedente)%> h/mês</b></td>
					<td colspan="3"><img src="imagens/blackdot.gif" width=1 height=0></td>
				</tr>
				<tr>
					<td colspan = "100%">
					cd_excedente: <%=excedente%> min - (<%=(total_mes(excedente))%>hs)<br>
					cd_minutos_trab: <%=minutos_trab%> min - (<%=(total_mes(minutos_trab))%>hs)<br>
					total_geral_trabalhado: <%=strsun%> min - (<%=(total_mes(strsun))%> hs)
				
					</td>
				</tr>
					<%End if%>
			
			</table>
		</td>
	</tr>
</table>
<br><br>

<SCRIPT LANGUAGE="javascript">

function JsInsert(cod_a1, cod_a2, cod_a3, cod_a4)
{
		document.location.href('empresa_funcionarios_obs.asp?dia_sel='+cod_a1+'&mes_sel='+cod_a2+'&ano_sel='+cod_a3+'&cd_codigo='+cod_a4+'&acao=inserir');
}
function JsEdit(cod_a, cod_b, cod_c, cod_d)
{
		document.location.href('empresa_funcionarios_obs.asp?dia_sel='+cod_a+'&mes_sel='+cod_b+'&ano_sel='+cod_c+'&cd_codigo='+cod_d+'&acao=editar');
}
function JsDelete(cod1, cod2, cod3, cod4)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('empresa/horas_trabalhadas/horas_trab_cartao_acao.asp?cd_codigo='+cod1+'&cd_cod_horas='+cod2+'&acao=excluir&mesentrada=<%=mes_entrada%>&anoentrada=<%=ano_entrada%>&menos_min='+cod3+'&cd_excedente='+cod4+'');
	  }
}
</SCRIPT>


</body>
</html>
