
<html>
<%	'Option Explicit
	Response.Buffer = True
%>

<!--include file="WriteBarcode128B.asp"-->
<!--#include file="../../includes/util.asp"-->
<head>
	<title>teste</title>
<%
	'Dim TextString
	'TextString = request("protocolo"&)'"200035542"
	
	unidade = request("unidade")
	nm_unidade = request("nm_unidade")
	'protocolo = request("proto_inicio")+1
	protocolo = "121"
	qtd_etiquetas = request("qtd_etiquetas")'*65
	
	imprimir = request("imprimir")
%>

<style type="text/css" media="screen"> 
	.etiquetas{
		font-weight: bold;
		font-size: 8px;
		color: #ff3300;
		font-family: verdana, Arial, Helvetica, sans-serif;
		}
	#no_print{ 
		visibility:visible; 
		display: block;}
	#ok_print{
		visibility:hidden; 
		display: none;}	
	
</style>
<style type="text/css" media="print"> 
	.etiquetas{
		font-weight: bold;
		font-size: 8px;
		color: #ff3300;
		font-family: verdana, Arial, Helvetica, sans-serif;
		}
	#ok_print{ 
		visibility:visible; 
		display: block;}
	#no_print{
		visibility:hidden; 
		display: none;}	
	
</style>
<script type="text/javascript">
	function printpage() {
	self.print();
	self.close()
	}
</script>
</head> 

<%if imprimir = 1 then%>
<body onload="printpage">
<%else%>
<body>
<%end if

'unidade = "20"
'nm_unidade = ".Hosp. Edmundo vasconcelos"
'protocolo = request("proto_inicio")



if unidade <> "" AND protocolo <> "" Then
	etapa = 1
end if

'TextString = unidade&protocolo&digito

'**************** etapa 1 # *********************
if etapa = 1 then
'************************************************

 coluna = "1"
 linha = "1"
 numero = "0"
 'msup = "22" 'Borda superior da etiqueta
 msup = "780"
 mesq =  "4" 'Borda esquerda da etiqueta
 
 msup = msup + 6 'posi��o do texto (superior)
 mesq = mesq 'posi��o do texto (esquerda). O texto j� est� centralizado
 
 
' qtd_etiquetas = "65"



while not numero = int(qtd_etiquetas)


%>
<div style="border:0px solid red; width:146px; height:81px; position:absolute; top:<%=msup%>px; left:<%=mesq%>px;">

<table cellspacing="2" cellpadding="2" border="0"  style="position:absolute; top:<%=msup-805%>px; left:<%=mesq+100%>px;">
	<tr>		
		<td><img src="../../imagens/Vdlap2p.gif" alt="" width="75" height="43" border="0"></td>
		<td>PROTOCOLO DE CONTROLE T�CNICO</td>
	</tr>
</table>

<p ></p>

<img src="../../imagens/blackdot.gif" alt="" width="500" height="1" border="0" style="position:absolute; top:<%=msup-750%>px; left:<%=mesq+30%>px;"> 
<img src="../../imagens/blackdot.gif" alt="" width="500" height="1" border="0" style="position:absolute; top:<%=msup-750%>px; left:<%=mesq+30%>px;">
<p align="center"  id="oks_print"><!--14 - coluna: <%'=coluna%><br>linha: <%'=linha%><br> numero:<%'=numero%>-->
<!--font color="#000000" face="Arial" size="2">VDLAP</font--><br>
<font color="#000000" face="Arial" size="1"><%=nm_unidade%></font><br>
<%' Call WriteBarcode128B(TextString) %><br>
<font color="#000000" face="Arial" size="2"><%=digitov(zero(unidade)&"."&proto(protocolo))%></font>
</p>
</div>
<%	if coluna = 5 then
		coluna = 0
		linha = linha + 1
		mesq = 4
		
			if linha = 14 then
				msup = msup + 135
				linha = 1
			else
				msup = msup + 80
			end if
	else
		msup = msup '+ 100
		mesq = mesq + 154
	end if
	
coluna = coluna + 1
numero = numero + 1
protocolo = protocolo + 1

wend%>
<br>
<%end if%>
<table border="1" id="nos_print" align="center">
	<form action="../../protocolo.asp" name="etiquetas" id="etiquetas">
	<input type="hidden" name="tipo" value="etiquetas">
	<input type="hidden" name="protocolo" value="<%=protocolo%>">
	<tr>
		<td colspan="2" align="center">Impress�o de etiquetas para protocolo</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<td>Unidade</td>
		
		<td><select name="unidade">
				<option value=""></option>
				<%strsql ="Select * from TBL_unidades order by A053_desung"
					Set rs_etq = dbconn.execute(strsql)%>
					<%Do while Not rs_etq.EOF%>
					<%cd_conv = rs_etq("A053_codung")%>
					<%if cd_etiqueta=cd_conv Then%>
					<%ck_etq="selected"%><%else ck_etq=""%>
					<%end if%>
					<option value="<%=rs_etq("A053_codung")%>" <%=ck_etq%>><%=rs_etq("A053_desung")%></option><%rs_etq.movenext
					Loop%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Qtd. folhas de Etiquetas</td>
		<td><select name="qtd_etiquetas">
				<%for i = 1 to 100%>
				<option value="<%=i%>"><%=zero(i)%></option>
				<%next%>
			</select>
		</td>
	</tr>
	<tr><%strsql ="Select * from TBL_unidades order by A053_desung"
					Set rs_etq = dbconn.execute(strsql)%>
		<td colspan="2">Ultima etiqueta impressa: <%=unidade%>.<%=protocolo%>-<%=digito%></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" name="Imprimir" value="1" width="20"><br>
		<!--input onclick="javascript:window.open('cooperativa/cooperados/coop_cadastro_impressao.asp?cod=<%=strcod%>','Impressao','width=700,height=185');" type="button" value="Imprimir" title="Apagar"--></td>
	</tr>
	</form>
</table>

</body>
