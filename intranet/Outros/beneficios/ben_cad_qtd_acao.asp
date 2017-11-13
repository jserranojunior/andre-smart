<!--#include file="../includes/inc_open_connection.asp"-->

<%

cd_codigo = request("cd_codigo")
cd_cod_ben = request("cd_cod_ben")
acao = request("acao")
mes_referencia = request("mes_referencia")
ano_referencia = request("ano_referencia")

cd_transp  =  request("cd_transp")
cd_ref  =  request("cd_ref")
cd_conv = request("cd_conv")
cd_inss = request("cd_inss")

cd_unidades = request("cd_unidades")
cd_funcao = request("cd_funcao")
cd_status = request("cd_status")

ordem = request("ordem")

If strstatus = "4" Then
strdt_demissao = FormatDateTime(now,2)

Else
strdt_demissao = "0"
End If

		if cd_status = 3 And acao = "inserir" Then
		'registra hora neutra na tabela de horas trabalhadas do cooperado no mes referente
			
			data_neutra = mes_referencia&"/01/"&ano_referencia&" 00:00"
			
			'xsql ="up_cartao_horas_insert @cd_codigo="&cd_codigo&",@data_entrada='"&data_entrada&"', @data_saida='"&data_saida&"',@cd_status="&cd_status&",@cd_unidades='"&cd_unidades&"',@cd_funcao='"&cd_funcao&"'"
			xsql ="up_cartao_horas_insert @cd_codigo="&cd_codigo&",@data_entrada='"&data_neutra&"', @data_saida_interv='"&data_neutra&"',@data_entr_interv='"&data_neutra&"',@data_saida='"&data_neutra&"',@cd_status="&cd_status&",@cd_unidades='"&cd_unidades&"',@cd_funcao='"&cd_funcao&"'"
			dbconn.execute(xsql)
			response.write("ok")
		
		Elseif cd_status = 3 And acao = "alterar" Then
		'Não faz nada por enquanto
		
		response.write("nada")
		Else
		response.write("Não")
		end if


If acao = "inserir" Then
    xsql = "up_Beneficios_calc_insert @cd_funcionario='"&cd_codigo&"', @nr_qtd_trans='"&cd_transp&"', @nr_qtd_ref='"&cd_ref&"', @nr_qtd_conv='"&cd_conv&"', @nr_qtd_inss='"&cd_inss&"', @dt_mes='"&mes_referencia&"', @dt_ano='"&ano_referencia&"'"
	dbconn.execute(xsql)
	strmsg = "Benefício cadastrado com sucesso!"
	response.redirect "../beneficios.asp?mes_sel="&mes_referencia&"&ano_sel="&ano_referencia&"&ordem="&ordem&""

ElseIf acao = "excluir" Then
    xsql= "up_Funcionario_delete @cd_funcionario='"&strcod&"'"
	dbconn.execute(xsql)
	strmsg = "Produto excluido com sucesso!!!"


ElseIf acao = "alterar" Then
    xsql = "up_Beneficios_calc_update @cd_cod_ben='"&cd_cod_ben&"', @nr_qtd_trans='"&cd_transp&"', @nr_qtd_ref='"&cd_ref&"', @nr_qtd_conv='"&cd_conv&"', @nr_qtd_inss='"&cd_inss&"'"
	dbconn.execute(xsql)
	strmsg = "Benefício alterado com sucesso!"
	response.redirect "../beneficios.asp?mes_sel="&mes_referencia&"&ano_sel="&ano_referencia&"&ordem="&ordem&""


Else
response.write("erro")

End if


Set Upload = Nothing 


'response.redirect "http://www.vdlap.com.br/adm23pto/beneficios_lista.asp?pag=1&tipo=nm_status&mes="&strmes&"&msg="&strmsg&""
'response.redirect "beneficios_lista.asp?pag=1&tipo=nm_status&mes="&strmes&"&msg="&strmsg&""


response.write(cd_codigo)&"codigo <br>"
response.write(cd_cod_ben)&"<br>"
response.write(acao)&"<br>"
response.write(cd_transp)&"<br>"
response.write(cd_ref)&"<br>"
response.write(cd_conv)&"<br>"
response.write(cd_inss)&"<br>"
response.write(mes_referencia)&"/"
response.write(ano_referencia)&"<br>"
response.write(cd_unidades)&"unidade<br>"
response.write(cd_funcao)&"funcao<br>"
response.write(cd_status)&"status<br>"

response.write(strmsg)
%>

   