
<html>
<%	'Option Explicit
	Response.Buffer = True
%>

<!--#include file="WriteBarcode128B.asp"-->
<head>
	<title>Etiquetas</title>

<%
	'Dim TextString
	TextString = "200035542"
%>
<style type="text/css" media="screen"> 
	.etiquetas{
		font-weight: bold;
		font-size: 8px;
		color: #ff3300;
		font-family: verdana, Arial, Helvetica, sans-serif;
		}	
	
</style>

</head> 
<body>
<%
unidade = "20"
nm_unidade = "Hosp. Edmundo vasconcelos"
prot_inicial = "3554"
protocolo = (prot_inicial)
digito = "2"

TextString = unidade&protocolo&digito


 coluna = "1"
 linha = "1"
 numero = "0"
 msup = "22" 'Borda superior da etiqueta
 mesq =  "4" 'Borda esquerda da etiqueta
 
 msup = msup + 6 'posição do texto (superior)
 mesq = mesq 'posição do texto (esquerda). O texto já está centralizado
 
 
 qtd_etiquetas = "195"

while not numero = int(qtd_etiquetas)%>
<div style="border:0px solid red; width:146px; height:81px; position:absolute; top:<%=msup%>px; left:<%=mesq%>px;"> 
<p align="center" class="etiquetas_"><!--14 - coluna: <%'=coluna%><br>linha: <%'=linha%><br> numero:<%'=numero%>-->
<font color="#000000" face="Arial" size="2">VDLAP</font><br>
<font color="#000000" face="Arial" size="1"><%=nm_unidade%></font><br>
<% Call WriteBarcode128B(TextString) %><br>
<font color="#000000" face="Arial" size="2"><%=mid(TextString,1,2)&"."&mid(TextString,3,4)&"-"&mid(TextString,9,1)%></font>
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

wend%>

</body>
