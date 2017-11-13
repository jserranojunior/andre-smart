<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%
variavel = request("cod_pat_manut")
'response.write(variavel)
sep_1 = instr(1,variavel,"#",1)
'*** cod. Patrimonio ***
cd_patrimonio = mid(variavel,1,sep_1)
cd_patrimonio = replace(cd_patrimonio,"#","")

variavel = mid(variavel,sep_1+1,len(variavel))
sep_2 = instr(1,variavel,"$",1)
'*** Tipo manutencao ***
cd_manutencao = mid(variavel,1,sep_2)
cd_manutencao = replace(cd_manutencao,"$","")

variavel = mid(variavel,sep_2+1,len(variavel))
sep_3 = instr(1,variavel,"*",1)
'*** Categoria ***
cd_categoria = mid(variavel,1,sep_2)
cd_categoria = replace(cd_categoria,"*","")

variavel = mid(variavel,sep_3+1,len(variavel))
sep_4 = instr(1,variavel,"|",1)
'*** Categoria ***
cod_confirma = mid(variavel,1,sep_2)
cod_confirma = replace(cod_confirma,"|","")
'response.write(cod_confirma)

'*** Data ***
dt_data = mid(variavel,sep_4+1,len(variavel))
dt_data = month(dt_data)&"/"&day(dt_data)&"/"&year(dt_data)


	'*** Verifica se o patrimonio existe na tabela de confirmações ***
	xsql ="SELECT * FROM TBL_Patrimonio_manutencoes_confirma WHERE cd_patrimonio='"&cd_patrimonio&"' AND cd_manutencao="&cd_manutencao&" AND YEAR(dt_data)='"&year(dt_data)&"' ORDER BY dt_data DESC"
	Set rs_confirm = dbconn.execute(xsql)
		'if not rs_confirm.EOF then
		
		'*** Se o cod de confirmação for 1, confirma a manutenção ***
		if int(cod_confirma) = 1 then
			if rs_confirm.EOF then
				xsql = "up_patrimonio_manut_confirma @cd_patrimonio='"&cd_patrimonio&"', @cd_manutencao='"&cd_manutencao&"', @cd_categoria='"&cd_categoria&"', @dt_data='"&dt_data&"'"
				dbconn.execute(xsql)
			end if
		
		'*** Se o cod de confirmação for 2, apaga a confirmação ***
		elseif int(cod_confirma) = 2 then
			if not rs_confirm.EOF then
			cd_confirma = rs_confirm("cd_codigo")
			
				xsql = "up_patrimonio_manut_desconfirma @cd_patrimonio='"&cd_confirma&"'"
				dbconn.execute(xsql)
				'response.write(cd_confirma)
			end if
		end if
	
	
	%>