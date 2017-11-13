<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--include file="../includes/inc_open_connection.asp"-->
<!--include file="../includes/util.asp"-->
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
'strano = 2007
formatahora = NovaHora

End Function


Function total_mes(min)

Hora = Int(min/60)   '60 minutos
Minutos = min Mod 60 '60 segundos
NovaHora = Right("0" & Hora,3) & ":" & Right("0" & Minutos,2)
'strano = 2007

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
document.form.diaentrada.focus();}


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

<body onLoad="foco()"><br>

<table width="650" cellspacing="0" cellpadding="0" border="1" class="textoPadrao" align="center">
	<tr>
		<td align="center">
			<table width="650" cellspacing="1" cellpadding="1" border="0" class="textoPadrao">
					<tr>
						<td colspan="100%"><b>FUNCIONÁRIOS - <font color="red">Horas Trabalhadas</font></b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					
				<%'*********************************************************
				  '*				      1ª Parte	 					  '*
				  '*********************************************************
					if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then%>
					
					<tr>
					    <td>Selecione um período abaixo:</td>
					    <td>&nbsp;</td>
						<td><!--ou escolha meses anteriores:-->&nbsp;</td>
					</tr>
					<tr align="center">
					    <td><%
						if mes_prod = "1" Then
						mes_ant = 12
						ano_ant = ano_prod - 1
						Else 
						mes_ant = mes_prod - 1
						ano_ant = ano_prod
						end if
						%>
						<a href=empresa_funcionarios_horas.asp?mes_sel=<%=mes_ant%>&ano_sel=<%=ano_ant%>><%=mes_selecionado(mes_ant)%>/<%=ano_ant%></a></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr align="center">
					    <td><a href=empresa_funcionarios_horas.asp?mes_sel=<%=mes_prod%>&ano_sel=<%=ano_prod%>><b><%=mes_selecionado(mes_prod)%>/<%=ano_prod%></b></a></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr align="center">
					    <td><%
						mes_prox = month(now)+ 1
						ano_prox = year(now)
						
						if mes_prox = "13" Then
						mes_prox = 1
						ano_prox = ano_prox + 1
						end if
						%>
						<a href=empresa_funcionarios_horas.asp?mes_sel=<%=mes_prox%>&ano_sel=<%=ano_prox%>><%=mes_selecionado(mes_prox)%>/<%=ano_prox%></a>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						</td>
					</tr>
					<tr>
					    <td>&nbsp;</td>
					    <td>&nbsp;</td>
					    <td><!--Selecione a data: <select name="Teste" id="Teste" size="">-->&nbsp;
					<td>
					</tr>
				<%'*********************************************************
				  '*				      2ª Parte	 					  '*
				  '*********************************************************
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
					
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="4">[<a href="empresa_funcionarios_horas.asp">Mudar Mês</a>] &nbsp;&nbsp; <b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b></td>
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
					<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('empresa_funcionarios_horas.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&cd_codigo=<%=cd_codigo%>');" onmouseout="mOut(this,'');">
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
						<td colspan=2>&nbsp;<%=conta_coop%> Cooperados</td></tr>
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
						<td colspan="5">[<a href="empresa_funcionarios_horas.asp">Mudar Mês</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp;[<a href="empresa_funcionarios_horas.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>">Mudar Cooperado]</td>
					</tr>
					
					<tr bgcolor=#cococo>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=235 height=1></td>
						<td><img src="imagens/blackdot.gif" width=5 height=1></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=235 height=1></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=1></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=1></td>
					</tr>
					<tr>
						<td colspan="100%"><br>&nbsp;&nbsp;<b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=235 height=0></td>
						<td><img src="imagens/blackdot.gif" width=5 height=0></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=235 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=0></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center width=244 colspan="2"> Entrada </td>
						<td>&nbsp;</td>
						<td align=center width=280 colspan="2"> Saida </td>
						<td align=center width=200> Total hora </td>
						<td align=center width=150> Ação </td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=235 height=1></td>
						<td><img src="imagens/blackdot.gif" width=5 height=1></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=235 height=1></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=1></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=1></td>
					</tr>
					
				</table>
				<table width="650" cellspacing="1" cellpadding="1" border="0" class="textoPadrao">
					<%
					'For i= &strmes_sel&
					
					'mes_1 = strmes_sel - 1
					'mes_2 = strmes_sel 
						
					'For i= mes_1 To mes_2 
					'For i= 0 To 2
					'  response.write "<tr><td>&nbsp;"
					'response.write "</td></tr>"
					
					
					'xsql = "up_Hora_Trabakhada_por_cd_funcionario @cd_funcionario="&strcod&",@mes="&i&",@ano='"&strano&"'"
					xsql2 = "up_Hora_Trabakhada_por_cd_funcionario @cd_funcionario="&cd_codigo&",@mes="&mes_sel&",@ano='"&ano_sel&"'"
					Set rs2 = dbconn.execute(xsql2)
					
					'response.write "<tr><td><a href=Funcionarios_cartao.asp?pag=1&cod='"&strcod&"'&tipo=nm_status&busca=""&mes_sel='"&strmes_sel&"'>"
					'mesdoano(i)
					'response.write "</a></td></tr>"
					
					
					'strsun = 0
					'Strdias = 0
					
					Do While Not rs2.eof 
					
					'strst_ = rs("nm_status")
					cd_cod_horas = rs2("cd_codigo")
					
					min = rs2("min")
					negativo = Instr(1,min,"-",1)'Compara se o caracter "-" aparece na string min
					If negativo = "1"Then
					strst_ = "0" 
					strsinal = "ERRO!"
					end if
					
					IF strst_ = "" Then
						strcor = "#eaeaea"
					ElseIf strst_ ="5"Then
						strcor = "#E0FFFF"
					Elseif strst_ = "0" Then
						strcor = "#ff8080"
					End If 
					
					%>
					<tr bgcolor="<%=strcor%>"  onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'<%=strcor%>');">
						<td align=center width=172><%=Day(rs2("data_entrada"))%>-<%=month(rs2("data_entrada"))%>-<%=year(rs2("data_entrada"))%></td>
						<td align=center width=172><%=hour(rs2("data_entrada"))%>:<%=minute(rs2("data_entrada"))%>:<%=second(rs2("data_entrada"))%></td>
						<td align=center width=5>&nbsp;<img src="../imagens/px.gif" width="1" height="20"></td>
						<td align=center width=172><%=Day(rs2("data_saida"))%>-<%=month(rs2("data_saida"))%>-<%=year(rs2("data_saida"))%></td>
						<td align=center width=172><%=hour(rs2("data_saida"))%>:<%=minute(rs2("data_saida"))%>:<%=second(rs2("data_saida"))%></td>
						<td width=200 align=center><%if strsinal = "" Then%><%=formatahora(rs2("min"))%><%else%><%=strsinal%><%end if%></td>
					<%If fechado = "1" Then%>
					    <Td width=200 align=center><input type="button" value="   ...   "></td>
					<%Else%>	
					
					
						<td width=200 align=center>
						<input onclick="javascript:JsInsert('<%=month(rs2("data_entrada"))%>','<%=year(rs2("data_entrada"))%>','<%=cd_codigo%>')" type="button" value="I" title="Incluir"> 
						<input onclick="javascript:JsEdit('<%=Day(rs2("data_entrada"))%>','<%=month(rs2("data_entrada"))%>','<%=year(rs2("data_entrada"))%>','<%=cd_codigo%>')" type="button" value="E" title="Editar"> 
						<input onclick="javascript:JsDelete('<%=cd_codigo%>','<%=cd_cod_horas%>')" type="button" value="A<%'=cd_cod_horas%>" title="Apagar"> 
						</td>
					<%end if%>
					</tr>
					
					<%
					
					strsun = strsun + rs2("min") 
					strdias = strdias + 1
					conta_dias = conta_dias + 1
					rs2.movenext
					strsinal = ""
					strst_ = ""
					Loop
					%>
				<tr bgcolor=#cococo>
						<td align=center colspan="2"><img src="imagens/px.gif" width=235 height=0></td>
						<td><img src="imagens/px.gif" width=5 height=0></td>
						<td align=center colspan="2"><img src="imagens/px.gif" width=235 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=100 height=0></td>
						<td align=center><img src="imagens/px.gif" width=100 height=0></td>
				</tr>
				<tr bgcolor="<%=strcor%>">
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td align="center"><b><%=total_mes(strsun)%></b></td>
					<td>&nbsp;</td>
				</tr>
				<%If fechado <> "1" Then%>
				<tr>
					<td align=center>Data Entrada</td>
					<td align=center >Hora Entrada</td>
					<td width=5 align=center></td>
					<td align=center>Data Saida</td>
					<td align=center>Hora Saida</td>
					<td></td></tr>
				<tr>
			   <%if acao = "editar" then
					xsql3 = "up_Hora_Trabalhada_lista @cd_funcionario="&cd_codigo&",@dia="&dia_sel&",@mes="&mes_sel&",@ano='"&ano_sel&"'"
					
						Set rs3 = dbconn.execute(xsql3)
						if NOT rs3.EOF Then
						cd_cod_horas = rs3("cd_codigo")
						
						dia_entrada = day(rs3("data_entrada"))
						mes_entrada = month(rs3("data_entrada"))
						ano_entrada = year(rs3("data_entrada"))
						hora_entrada = hour(rs3("data_entrada"))
						minuto_entrada = minute(rs3("data_entrada"))
						
						dia_saida = day(rs3("data_saida"))
						mes_saida = month(rs3("data_saida"))
						ano_saida = year(rs3("data_saida"))
						hora_saida = hour(rs3("data_saida"))
						minuto_saida = minute(rs3("data_saida"))
						
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
			 
					<td align=center class="txt_cinza"><input type="txt" name="diaentrada" size=2 maxlength="2" value="<%=dia_entrada%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<%=mes_entrada%>/<%=ano_entrada%></td>
					<td align=center ><input type="txt" name="horaentrada_hr" size=2 maxlength="2" value="<%=hora_entrada%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="txt" name="horaentrada_min" size=2 maxlength="2" value="<%=minuto_entrada%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center></td>
				  	<td align=center><input type="txt" name="diasaida" size=2 maxlength="2" value="<%=dia_saida%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="txt" name="messaida" size=2 maxlength="2" value="<%=mes_saida%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="txt" name="anosaida"  size=4 maxlength="4" onKeyPress="ChecarTAB();"  value="<%=ano_saida%>" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
					<td align=center><input type="txt" name="horasaida_hr" size=2 maxlength="2" value="<%=hora_saida%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">:<input type="txt" name="horasaida_min" size=2 maxlength="2" value="<%=minuto_saida%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
					<td align=center><input type="submit" value="<%=acao%>"></td>
				</tr>
				<%else%>
				<tr>
					<td colspan="100%">&nbsp;<br><br></td>
				</tr>
				<%end if%>
				<tr>
					<td colspan="100%"><br>&nbsp; Total de dias trabalhados: <b><%=conta_dias%></b></td>
				</tr>
			
					<%End if%>
			
			</table>
		</td>
	</tr>
</table>
<br><br>

<SCRIPT LANGUAGE="javascript">

function JsInsert(cod_a1, cod_a2, cod_a3)
{
		document.location.href('empresa_funcionarios_horas.asp?mes_sel='+cod_a1+'&ano_sel='+cod_a2+'&cd_codigo='+cod_a3+'&acao=inserir');
}
function JsEdit(cod_a, cod_b, cod_c, cod_d)
{
		document.location.href('empresa_funcionarios_horas.asp?dia_sel='+cod_a+'&mes_sel='+cod_b+'&ano_sel='+cod_c+'&cd_codigo='+cod_d+'&acao=editar');
}
function JsDelete(cod1, cod2)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('empresa/horas_trabalhadas/horas_trab_cartao_acao.asp?cd_codigo='+cod1+'&cd_cod_horas='+cod2+'&acao=excluir&mesentrada=<%=mes_entrada%>&anoentrada=<%=ano_entrada%>');
	  }
}
</SCRIPT>
</body>
</html>



