<!--#include file="../../includes/util.asp"-->
<%
'sequencia = unidade&proto(protocolo)
unidade = "3"
protocolo = "121"
'mult = 2
qtd = 0
linha = 1
'protocolo = zero(unidade)&"."&proto(protocolo)
cor = "black"
%>


<% do while not qtd = 1000
proto_final = digitov(zero(unidade)&proto(protocolo))

qual_digito = mid(proto_final,10,1)

if qual_digito = ultimo then
cor = "red"
else
cor = "black"
End if



%>
<font color="<%=cor%>"><%=proto_final%> </font> &nbsp;&nbsp;&nbsp;<%if linha = 5 then%><br><br> <%linha = 0%><%end if%>

<%'if qtd = 0 then 
protocolo = protocolo + 1 
'end if
qtd = qtd +1
linha = linha + 1
ultimo = qual_digito
loop

%>