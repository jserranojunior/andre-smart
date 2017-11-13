
  <!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
   <!--#include file="../../includes/inc_area_restrita.asp"-->

<%data_atual = Month(now)&"/"&Day(now)&"/"&Year(now)&" "&Hour(now)&":"&Minute(now)&":"&Second(now)
cd_user = session("cd_codigo")

acao = request("acao")
filtro = request("filtro")
jan = request("jan")

sequencia = request("sequencia")
cd_codigo = request("cd_codigo")
campo = request("campo")

cd_orcamento = request("cd_orcamento")
nm_valor_os = request("nm_valor_os")
nm_valor_orc = request("nm_valor_orc")
nr_parcela = request("nr_parcela")

		if nr_parcela = "" OR nr_parcela = 1 Then 
			cd_tipo_valor = 1
		else
			cd_tipo_valor = 3
		end if
dt_orcamento = request("dt_orcamento")
dt_aprov_orc = request("dt_aprov_orc")
	'IF NOT ISNULL(dt_aprov_orc) THEN
	IF dt_aprov_orc <> "" THEN
		dt_dia_aprov = day(dt_aprov_orc) 'request("dt_dia_aprov")
		dt_mes_aprov = month(dt_aprov_orc) 'request("dt_mes_aprov")
		dt_ano_aprov = year(dt_aprov_orc) 'request("dt_ano_aprov")
			If dt_dia_aprov = "" AND dt_mes_aprov = "" AND  dt_ano_aprov = "" Then
				dt_orcamento_aprov = "Null"
			else
				'dt_orcamento_aprov = "'"&dt_mes_aprov&"/"&dt_dia_aprov&"/"&dt_ano_aprov&"'"
				dt_orcamento_aprov = "'"&dt_dia_aprov&"/"&dt_mes_aprov&"/"&dt_ano_aprov&"'"
			End if
			
			'**** Data Inicial ****
			dt_ano_inicial = dt_ano_aprov
			dt_mes_inicial = dt_mes_aprov+1
				if dt_mes_inicial > 12 Then
					dt_mes_inicial = 1
					dt_ano_inicial = dt_ano_inicial + 1
				end if
			dt_dia_inicial = 1
			dt_inicial = dt_ano_inicial&"-"&dt_mes_inicial&"/"&dt_dia_inicial
			
			'**** Data Final ****
			dt_ano_final = dt_ano_aprov
			dt_mes_final = dt_mes_aprov 
			for i=1 to nr_parcela
				
				dt_mes_final = dt_mes_final + 1
					if dt_mes_final > 12 Then
					dt_mes_final = 1
					dt_ano_final = dt_ano_final + 1
					
					end if
					'response.write("<br>"&dt_mes_final&"/"&dt_ano_final&"<br>")
					dt_dia_final = ultimodiames(dt_mes_final,dt_ano_final)
			next
			dt_final = dt_ano_final&"-"&dt_mes_final&"-"&dt_dia_final
		
		ELSE
			dt_inicial = year(dt_orcamento)&"-"&month(dt_orcamento)&"-1"
			dt_final = NULL
		'x="a"
		END IF



x = ""



'*******************************************************************************
'    					Vincula - Orçamentos									
'*******************************************************************************
if acao = "1" then
	'*** verifica se já está cadastrado ****
	xsql = "SELECT cd_codigo FROM TBL_manutencao_orc_aprov WHERE cd_orcamento="&cd_orcamento
	SET rs = dbconn.execute(xsql)
	if rs.EOF Then
		xsql = "INSERT INTO TBL_manutencao_orc_aprov (cd_orcamento,dt_inicial,dt_final,cd_tipo_valor,cd_user,dt_cad) VALUES ("&cd_orcamento&",'"&dt_inicial&"','"&dt_final&"',"&cd_tipo_valor&","&cd_user&",'"&data_atual&"')"
		dbconn.execute(xsql)
	end if

Elseif acao = "3" Then
	xsql = "up_os_delete2 @cd_os='"&cd_codigo_os&"'"
	dbconn.execute(xsql)
	

Else
action = "erro - teste"

End If%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>Untitled</title>
</head>

<body onload="window.close();" onunload="window.opener.location.reload(true);">
<br>
*** Movimentação ***<br>
jan: <%=jan%><br>
<br><br>


<br>


<br>

link: <a href ="<%=link%>">Redireciona</a>

</body>
</html>
