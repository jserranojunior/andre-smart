<!--#include file="../../includes/util.asp"-->
<%'sequencia = unidade&proto(protocolo)
unidade = "3"
sequencia = "120"
mult = 2
do while not qtd = 10
	
	'******************************
	'Calculo do digito verificador
	'******************************
	sequencia = proto(sequencia)
	protocolo = zero(unidade)&sequencia
	
	for i=1 to len(protocolo)
		asd = mid(protocolo,i,1)
		
		'prot = asd&" "&prot
		'protm = mult&" "&protm
		'res = (asd*mult)&" "&res
		res_som = (asd*mult)+res_som
		
		mult = mult + 1
		
	next
	'mult = 2
	oprdv = res_som/11
	oprdv1 = round(oprdv,1)
	resto = instr(oprdv1,",")
	resto1 = mid(oprdv1, resto+1, 1)
	dv = 11 - resto1
		'if dv = 0 or dv=1 or dv = 10 then
		'	dv = 0
		'end if
	'********************************

%>

(<%=oprdv1%>)
<%=zero(unidade)%>.<%=sequencia%>-<%=dv%><br><br>

<%sequencia = sequencia + 1 
qtd = qtd +1
loop%>