<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<%
dia_prod = day(now)
mes_prod = month(now)
ano_prod = year(now)

data_atual = dia_prod&"/"&mes_prod&"/"&ano_prod

dia_sel = request("dia_sel")
mes_sel = request("mes_sel")
ano_sel = request("ano_sel")

cd_codigo = request("cd_codigo")
ordem = request("ordem")

if ordem = "" then
ordem = "nm_nome,nm_sobrenome"
'ordem = "nm_unidade,nm_nome,nm_sobrenome"
end if

acao = request("acao")

	If mes_sel - 1 = 0 Then 
		mes_sel_1 = 12
		ano_sel_1 = ano_sel - 1
	Else
		mes_sel_1 = mes_sel - 1
		ano_sel_1 = ano_sel
	End if

xsql_verif = "up_Verifica_fechamento_mes @cd_mes='"&mes_sel_1&"', @cd_ano='"&ano_sel_1&"',@cd_origem=1"
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

<table width="500" cellspacing="0" cellpadding="0" border="1" class="textoPadrao" align="center">
	<tr>
		<td align="center">
			<table width="500" cellspacing="0" cellpadding="0" border="1" class="textoPadrao" bgcolor="#fbfbfb">
					<tr>
						<td colspan="100%">&nbsp; <b>BENEFÍCIOS - <font color="red">Cadastro de quantidades</font></b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					
<%'***************************
  '**		 1ª Parte		**
  '***************************
  
					if mes_sel = "" AND ano_sel = "" AND cd_codigo = "" Then%>
					
					<tr>
					    <td colspan="3">&nbsp; Selecione um período abaixo:</td>
					</tr>
					<!-- mes anterior -->
					<%mes_ant = mes_prod -1
						If mes_ant = 0 Then
						mes_ant = 12
						ano_ant = ano_prod - 1
						Else
						ano_ant = ano_prod
						End if%>
					
					<tr align="center">
					    <td colspan="3"><a href=beneficios.asp?mes_sel=<%=mes_ant%>&ano_sel=<%=ano_ant%>><%=mes_selecionado(mes_ant)%>/<%=ano_ant%></a><font color="#b3b3b3"> - referente a 15/<%=left(mes_selecionado(mes_ant),3)%> a
																																													<%if mes_ant +1 =13 then%>
																																													<%mes_ant = 1
																																													end if%>
																																													<%=left(mes_selecionado(mes_ant),3)%></font></td>
						
						
					</tr>
					<!-- mes atual -->
					<tr align="center">
					    <td colspan="3"><%
						mes_prox = month(now)+ 1
						ano_prox = year(now)
						
						if mes_prox = "13" Then
						mes_prox = 1
						ano_prox = ano_prox + 1
						end if
						%>
						<a href=beneficios.asp?mes_sel=<%=mes_prod%>&ano_sel=<%=ano_prod%>><b><%=mes_selecionado(mes_prod)%>/<%=ano_prod%></b></a><font color="#b3b3b3"> - referente: a 15/<%=left(mes_selecionado(mes_prod),3)%> a 
																																															<%if mes_prod = 12 then%>
																																															<%mes_ = 1
																																															Else
																																															mes_ = mes_prod + 1
																																															end if%>
																																															<%=left(mes_selecionado(mes_),3)%></font>		
						</td>
					</tr>
					<!-- mes seguinte -->
					<tr align="center">
					    <td colspan="3"><%
						mes_prox = month(now)+ 1
						ano_prox = year(now)
						
						if mes_prox = "13" Then
						mes_prox = 1
						ano_prox = ano_prox + 1
						end if
						%>
						<a href=beneficios.asp?mes_sel=<%=mes_prox%>&ano_sel=<%=ano_prox%>><%=mes_selecionado(mes_prox)%>/<%=ano_prox%></a><font color="#b3b3b3"> - referente: a 15/<%=left(mes_selecionado(mes_prox),3)%> a
																																													<%if mes_prod = 12 then%>
																																													<%mes_ = 1
																																													Else
																																													mes_ = mes_prod + 1
																																													end if%>
																																													<%=left(mes_selecionado(mes_),3)%></font>
						</td>
					</tr>
					<tr>
					    <td colspan="3">&nbsp;</td>
					
					</tr>
<%'***************************
  '**		 2ª Parte		**
  '***************************'

					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo = "" Then%>
					
					<tr>
						<td colspan="7">&nbsp; [<a href="beneficios.asp">Mudar Mês</a>] &nbsp;&nbsp; <b>( Pagamento em 15 de <%=mes_selecionado(mes_sel+1)%>/<%=ano_sel%>)</b></td>
					</tr>
					<tr bgcolor=#cococo>
						<td align=center colspan="100%"><img src="imagens/px.gif"  height=0></td>
					</tr>
					<tr>
						<td colspan="100%"><br>&nbsp;<b>Ordem:</b> 
						<a href="beneficios.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&ordem=nm_nome,nm_sobrenome">(Alfabética)</a>&nbsp;
						<a href="beneficios.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&ordem=nm_unidade,nm_nome,nm_sobrenome">(Unidade)</a>
						<br>&nbsp;</td>
					</tr>
					<tr bgcolor="#d7d7d7">
						<td><b>&nbsp;Cód.</b></td>
						<td><b>&nbsp;Matrícula</b></td>
						<td><b>&nbsp;Nome</a></b></td>
						<td><b>&nbsp;Tran</b></td>
						<td><b>&nbsp;Ref</b></td>
						<td><b>&nbsp;Conv</b></td>
						<td><b>&nbsp;Inss</b></td>
					</tr>
					
					<%
					conta_coop = "1"
					
					xsql = "up_funcionarios_lista @cd_codigo='', @ordem='"&ordem&"', @condicao=''"
					Set rs = dbconn.execute(xsql)
					
					Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					cd_unidades = rs("cd_unidades")
					dt_demissao = rs("dt_desliga")
					cd_status = rs("cd_status")
					nm_unidade = rs("nm_unidade")
					nm_sigla = rs("nm_sigla")
					
					
						'******* Verifica se já foi cadastrado benefício para o período ****
						xsql2 = "TBL_calc_beneficios Where cd_funcionario='"&cd_codigo&"' AND dt_mes='"&mes_sel&"' AND dt_ano='"&ano_sel&"'"
					
					Set rs_b = dbconn.execute(xsql2)
					
					If NOT rs_b.EOF Then
					trans_ver = rs_b("nr_qtd_trans")
					ref_ver = rs_b("nr_qtd_ref")
					conv_ver = rs_b("nr_qtd_conv")
					conv_ = conv_ver
					inss_ver = rs_b("nr_qtd_inss")
						if inss_ver = 1 Then
						inss_ver = "forçado"
						Else
						inss_ver = ""
						end if
					End if
					
					if cd_status = "4" Then
					cor_font = "#afafaf"
					'Else
					end if
					conta_coop = conta_coop + 1
					
					
					
					
					if instr(ordem,"unidade") <> 0 Then 
						if nm_unidade <> nm_unid then
							if nm_unid <> "" Then%>						
							<tr><td colspan="8">&nbsp;</td></tr>
							<%end if%>
						<tr><td colspan="8">&nbsp;<b><%=nm_unidade%></b></td></tr>
						<%end if%>
					<%End if%>
					<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:location.replace('beneficios.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>&cd_codigo=<%=cd_codigo%>&ordem=<%=ordem%>');" onmouseout="mOut(this,'');">
						<td><font color="<%=cor_font%>">&nbsp;<%=cd_codigo%></font></td>
						<td><font color="<%=cor_font%>">&nbsp;<%=cd_matricula%></font></td>
						<td><font color="<%=cor_font%>">&nbsp;<%=nm_nome%><%=nm_sobrenome%> <%'=cd_status%> <%'=dt_demissao%></font></td>
						<td><font color="<%=cor_font%>">&nbsp;<%=trans_ver%></font></td>
						<td><font color="<%=cor_font%>">&nbsp;<%=ref_ver%></font></td>
						<td><font color="<%=cor_font%>">&nbsp;<%=Conv_ver%></font></td>
						<td><font color="<%=cor_font%>">&nbsp;<%=inss_ver%></font></td>
					</tr>
					<%
					'end if
					trans_ver = ""
					ref_ver = ""
					conv_ver = ""
					inss_ver = ""
					dt_demissao = ""
					cor_font = ""
					
					
					nm_unid = nm_unidade
					
					rs.movenext
					Loop
					%>
					<tr><td colspan=7>&nbsp;</td></tr>
					<tr align="left">
						<td colspan=7>&nbsp;<%=conta_coop%> Cooperados</td></tr>
						
<%'***************************
  '**		 3ª Parte		**
  '***************************
				  
					elseif mes_sel <> "" AND ano_sel <> "" AND cd_codigo <> "" Then
					
				'******* Verifica a quantidade de convenios do cooperado ****
					if mes_sel = "1" then
					mes_conv = "12"
					ano_conv = ano_sel - 1
					Else
					mes_conv = mes_sel - 1
					ano_conv = ano_sel
					End if
					
					xsql1 = "TBL_calc_beneficios Where cd_funcionario='"&cd_codigo&"' AND dt_mes='"&mes_conv&"' AND dt_ano='"&ano_conv&"'"
					Set rs_conv = dbconn.execute(xsql1)
					if Not rs_conv.EOF Then
					conv_ant = rs_conv("nr_qtd_conv")
					'response.write("ok")
					End if
					
				'******* Verifica se já foi cadastrado benefício para o período ****
					xsql2 = "TBL_calc_beneficios Where cd_funcionario='"&cd_codigo&"' AND dt_mes='"&mes_sel&"' AND dt_ano='"&ano_sel&"'"
					Set rs_b = dbconn.execute(xsql2)
					
					If NOT rs_b.EOF Then
					cd_cod_ben = rs_b("cd_cod")
					trans_ver = rs_b("nr_qtd_trans")
					ref_ver = rs_b("nr_qtd_ref")
					conv_ver = rs_b("nr_qtd_conv")
					'conv_ = conv_ver
					inss_ver = rs_b("nr_qtd_inss")
					End if
					
					
					xsql = "up_funcionarios_lista @cd_codigo='"&cd_codigo&"', @condicao='', @ordem='nm_nome, nm_sobrenome'"
					Set rs = dbconn.execute(xsql)
					
					'Do while NOT rs.EOF
					
					cd_codigo = rs("cd_codigo")
					cd_matricula = rs("cd_matricula")
					nm_nome = rs("nm_nome")
					nm_sobrenome = rs("nm_sobrenome")
					nm_foto = rs("nm_foto")
					cd_unidades = rs("cd_unidades")
					dt_demissao = rs("dt_desliga")
					cd_status = rs("cd_status")
					'cd_funcao = rs("cd_funcao")
					
					
					
					
					if dt_demissao <= data_atual Then
					status = "(desligado)"
					end if
					%>
					<tr>
						<td colspan="4">&nbsp; [<a href="beneficios.asp">Mudar Mês</a>]  &nbsp;&nbsp;/  &nbsp;&nbsp;[<a href="beneficios.asp?mes_sel=<%=mes_sel%>&ano_sel=<%=ano_sel%>">Mudar Cooperado]</td>
					</tr>
					<tr><td align=center bgcolor=black colspan=5></td> </tr>
					<tr rowspan="2">
   						<td align=left valign="middle" class="textopadrao">
							<br><br>&nbsp; 
							<font color="red"><b><%=mes_selecionado(mes_sel)%>/<%=ano_sel%></b></font>
							<br><br>&nbsp; 
							<%=rs("nm_nome")%>&nbsp;<%=rs("nm_sobrenome")%>
							<br><br>&nbsp; 
							Transporte: <b><%=trans_ver%></b><br>&nbsp; 
							Refei&ccedil;&atilde;o: <b><%=ref_ver%></b><br>&nbsp; 
							Conv&ecirc;nio: <b><%=conv_ver%></b><br>&nbsp; 
							INSS: <%if inss_ver = "1" Then%><font color="#008000"><b><%response.write("Caso especial")%></b></font><%end if%><br><br>
							&nbsp; <font color="red">Status atual:</font><%=fechado%><%xsql_status = "SELECT * FROM TBL_Status Where cd_codigo = '"&cd_status&"'"
															Set rs_status = dbconn.execute(xsql_status)%> <b><%response.write(rs_status("nm_status"))%></b>
							<br><br><br>
						</td>
						
   						<td align=left class="textopadrao">
							<% If nm_foto <> "" then%>
							<img src="foto/funcionarios/<%=nm_foto%>" width=100 >
							<%Else%>
							<img src="imagens/Vdlap2p.gif" width=100>
							<br><br>
							<%End if%></td>
 					</tr>
					<%
					If trans_ver = "" OR ref_ver = "" OR conv_ver = "" Then
					acao = "inserir"
					Else
					acao = "alterar"
					End if
					%>
 <Tr bgcolor=#cococo class="textopadrao">
	<form name="form" action="beneficios/ben_cad_qtd_acao.asp" id="beneficios" method="post">
		<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">
		<input type="hidden" name="cd_cod_ben" value="<%=cd_cod_ben%>">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="mes_referencia" value="<%=mes_sel%>">
		<input type="hidden" name="ano_referencia" value="<%=ano_sel%>">
		<input type="hidden" name="ordem" value="<%=ordem%>">
		
		<input type="hidden" name="cd_unidades" value="<%=cd_unidades%>">
		<input type="hidden" name="cd_status" value="<%=cd_status%>">
		<input type="hidden" name="cd_funcao" value="<%=cd_funcao%>">
 </tr>
 <tr><td align=center bgcolor="#d6d6d6" colspan=5></td> </tr>
 <tr class="textopadrao"> 
 
<%If fechado = "1" Then%>
 <td align=left>&nbsp;
</td>
<%Else%>
		<td align=left colspan=2><br>
		<li>Vale Transporte ............<input type="txt" name="cd_transp" size=5 maxlength="3" value="<%=trans_ver%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 3)" onFocus="PararTAB(this);"><br>
 		<li>Vale Refei&ccedil;&atilde;o ...............<input type="txt" name="cd_ref" size=5 maxlength="2" value="<%=ref_ver%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"><br>
		<li>Conv&ecirc;nio .....................<input type="txt" name="cd_conv" size=5 maxlength="2" <%if acao = "inserir" Then%> value="<%=conv_ant%>" <%Elseif acao="alterar" Then%> value="<%=conv_ver%>" <%end if%> onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">
  		<%if inss_ver = "1" Then
		inss_check = "checked"
		End if
		%>
		<li>INSS ...........................<input type="checkbox" name="cd_inss" value="1" <%=inss_check%>>*<br>&nbsp; 
		

</td>
</tr><tr>
  <td align=Center colspan=3><br><input type="submit" value="OK"><!--input type="button" value="Voltar" onClick="Javascript:window.location='beneficios_lista.asp?mes=<%=mes_sel%>'"-->&nbsp;&nbsp;&nbsp;</td> 
 <%end If%>
   </form>
 
 </tr>
 <tr><td align=left colspan=3>&nbsp;</td></tr>
 
 <tr><td align=center bgcolor=black colspan=5></td> </tr>
 
 <tr><td colspan="100%" class="textopequeno">&nbsp; * somente marque INSS para pagamento em caso especial</td></tr>
 
 

				</table>
				
			
					<%End if%>

		</td>
	</tr>
</table>
<br><br>
</body>
</html>



