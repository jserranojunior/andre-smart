<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<%tela = request("tela")

if tela = "" Then
	div_height = "200px"
'Else
	'div_height = "auto"
end if
%>
<style type="text/css">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt {color: #000000;font-family: verdana;font-size:12px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:12px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:1px solid #000000; }  
.print {color: #000000;font-family: verdana;font-size:9px;text-decoration:none;}
.divrolagem {
/* define barra de rolagem automatica quando o 
conteudo ultrapassar o limite em x ou y */
    overflow: auto; 
/* define o limite maximo da autura do div */
    height: <%=div_height%>;
/* define o limite maximo da largura do div */
    width: auto;}
[code]

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
response.write("OK, Exclu�do")
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
	<TITLE> Relat�rio do Banco de Horas </TITLE>
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
<table width="600" cellspacing="0" cellpadding="0" border="0" class="textoPadrao" align="center">
	<tr>
		<td align="center">
			<table width="100%" cellspacing="1" cellpadding="0" border="0" class="txt" bordercolordark="#FFFFFF">
					<!--tr bgcolor=#cococo>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=150 height=0></td>
						<td align=center><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td align=center colspan="2"><img src="imagens/blackdot.gif" width=200 height=0></td>
						<td><img src="imagens/blackdot.gif" width=1 height=0></td>
						<td><img src="imagens/blackdot.gif" width=80 height=0></td>
						<td><img src="imagens/blackdot.gif" width=75 height=0></td>
					</tr-->
					<tr>
						<td colspan="100%">&nbsp;<b>FUNCION�RIOS - <font color="red">Banco de horas</font></b></td>
					</tr>
					<tr>
						<td align=center colspan="100%">&nbsp;<!--img src="../../imagens/px.gif" height="10"--></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="../../imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					<tr>
						<td>&nbsp;<b>Nome&nbsp;</b></td>
						<!--td align="right"><b>Mar&nbsp;</b></td>
						<td align="right"><b>Abr&nbsp;</b></td>
						<td align="right"><b>Mai&nbsp;</b></td>
						<td align="right"><b>Jun&nbsp;</b></td-->
						<td align="right"><b>Saldo&nbsp;</b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="../../imagens/blackdot.gif" width="100%" height=0></td>
					</tr>
					
					
										
						<%nm_lista = ""
						cd_meses = 1
						qtd_meses = 0
					
						
						'xsql = "select * from View_relatorio_banco_horas_meses"
						xsql = "select * from View_relat_banco_horas_total"
						Set rs = dbconn.execute(xsql)
						Do while NOT rs.EOF
					
						cd_codigo = rs("cd_codigo")
						nm_nome = rs("nome")
'						dt_mes = rs("dt_mes")
						'cd_saldo = rs("cd_horas_excedentes")
						cd_saldo = rs("h_total")
								'*** Procura saldo negativo e avisa ***
									str_negativo = Instr(1,cd_saldo,"-",1)'Compara se o caracter "-" aparece na string min
									If str_negativo = "1" Then
									cor_saldo = "red"
									Else
									cor_saldo  = "black"
									end if
								'****************************************************************************
						%>
						<%if nm_lista <> nm_nome Then%>
					<tr onmouseover="mOvr(this,'#CFC8FF');" onmouseout="mOut(this,'#F2F2F2');" bgcolor="#F2F2F2">
						<td>&nbsp;<%=nm_nome%></td>
						<%Else%>
						
						<%end if%>
						
						<!--td><%=dt_mes%></td-->
						<td align="right"><font color=<%=cor_saldo%>><%=total_mes(cd_saldo)%>&nbsp;</font></td>
						
						<%if cd_meses = qtd_meses then%>
							<%xsql2 = "select * from View_relatorio_banco_horas_total where cd_codigo='"&cd_codigo&"'"
							Set rs2 = dbconn.execute(xsql2)
							Do while NOT rs2.EOF
						
							saldo = rs2("saldo")
							
							
								'*** Procura saldo negativo e avisa ***
									str_negativo = Instr(1,saldo,"-",1)'Compara se o caracter "-" aparece na string min
									If str_negativo = "1" Then
									cor_saldo = "red" 
									Else
									cor_saldo  = "black"
									end if
								'****************************************************************************
							%>
							
						<td align="right"><font color=<%=cor_saldo%>><b><%=total_mes(saldo)%>&nbsp;</b></font></td>
							<% rs2.movenext
							loop%>
						<%end if%>
					
					<%rs.movenext
					nm_lista = nm_nome
					cd_meses = cd_meses + 1
					cor_saldo = ""
						if cd_meses > 3 then
							cd_meses = 1
						end if
					loop%>
					</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>