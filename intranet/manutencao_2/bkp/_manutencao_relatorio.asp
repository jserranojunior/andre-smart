<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<style type="text/css" media="screen">
<!--
.txt_titulo {color: #000000;font-family: verdana;font-size:16px;}
.txt_menu {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:hover {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:link {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_menu:visited {color: #FFFFFF;font-family: verdana;font-size:10px;}
.txt_ {color: #000000;font-family: verdana;font-size:11px;text-decoration:none;}
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:10px;text-decoration:none; }
.inputs { background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:0px solid #000000; }  
.divrolagem {
/* define barra de rolagem automatica quando o 
conteudo ultrapassar o limite em x ou y */
    overflow: auto; 
/* define o limite maximo da autura do div */
    height: <%=div_height%>;
/* define o limite maximo da largura do div */
    width: auto;}
#no_print{ 
	visibility:visible; 
	display: block;}
#ok_print{
	visibility:hidden; 
	display: none;}
table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	border-color: Gray Gray Gray Gray;
	font-family: verdana;
	font-size:11px;
	text-decoration:none;}
-->
</style>
<style type="text/css" media="print">
<!--
.txt_ {color: #000000;font-family: verdana;font-size:10px;text-decoration:none;}
#no_print{ 
	visibility:hidden; 
	display: none;}
#ok_print{
	visibility:visible;
	display:block;}
table{
	background-color: #ffffff; 
    border: 0px solid #cccccc;
	width: 100px;
	font-family: verdana;
	font-size:10px;
	text-decoration:none;}
.divrolagem {
	overflow: auto; 
    height: auto;
    width: auto;}
-->
</style>
<script language="javascript">
<!--
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
  
// function marca_todos()
// 	document.form.campos.value=h_data;
	
	
 function selecionar_tudo(){ 
   for (i=0;i<document.form.elements.length;i++)
      if(document.form.elements[i].type == "checkbox"){
        document.form.elements[i].checked=1}
	}
 function deselecionar_tudo(){ 
  for (i=0;i<document.form.elements.length;i++) 
     if(document.form.elements[i].type == "checkbox"){
         document.form.elements[i].checked=0}
	}  

// -->	
</script>


<%x = request("x")
ano_atual = Year(now)

'*** DETALHES ***************
questao = request("questao")
campos = request("campos")
cd_ns = request("cd_ns")
cd_patrimonio = request("cd_patrimonio")
cd_patrimonio_vazio = request("cd_patrimonio_vazio")
nm_patrimonio = request("nm_patrimonio")
cd_equipamento = request("cd_equipamento")
nm_equipamento = request("nm_equipamento")
nm_unidade = request("nm_unidade")
cd_fornecedor = request("cd_fornecedor")
cd_especialidade = request("cd_especialidade")
cd_avaliacao = request("cd_avaliacao")
cd_marca = request("cd_marca")
os_pendente = request("os_pendente")


'***** Pendentes ******
if os_pendente = "0" Then 'O.S. Pendente
cond_fecha = "AND (fecha BETWEEN 0 AND 0)"
ck_pend1 = "checked"
elseif os_pendente = "1" then ' O.S. Encerrada
cond_fecha = "AND (fecha BETWEEN 1 AND 1)"
ck_pend2 = "checked"
Elseif os_pendente = "" Then
cond_fecha = " AND (fecha BETWEEN '0' AND '1')"
ck_pend3 = "checked"
End if



'***** Verifica quais campos foram selecionados *****
	campo_0 = Instr(1,campos,"num_os",0)
	campo_2_1 = Instr(1,campos,"cd_ns",0)
	campo_2_2 = Instr(1,campos,"cd_patrimonio",0)
	campo_2_3 = Instr(1,campos,"nm_patrimonio",0)
	campo_2_4 = Instr(1,campos,"cd_patrimonio_vazio",0)
	campo_1 = Instr(1,campos,"nm_equipamento",0)
	campo_2 = Instr(1,campos,"nm_unidade",0)
	campo_3 = Instr(1,campos,"cd_fornecedor",0)
	campo_4 = Instr(1,campos,"cd_especialidade",0)
	campo_5 = Instr(1,campos,"nm_avaliacao",0)
	campo_6 = Instr(1,campos,"nm_marca",0)
	'campo_7 = Instr(1,campos,"dt_dia",0)
	campo_7 = Instr(1,campos,"dt_os",0)
	'campo_7 = Instr(1,campos,"dt_ano",0)
	campo_8 = Instr(1,campos,"num_os_assist",0)
	campo_12 = Instr(1,campos,"cd_valor_orc",0)


'***** Tratamento da ordem de listagem ******
ordem_conta = request("ordem_conta")
ordem_os = request("ordem_os")
ordem_ns = request("ordem_ns")
ordem_patrimonio = request("ordem_patrimonio")
ordem_equip = request("ordem_equip")
ordem_unidade = request("ordem_unidade")
ordem_fornecedor = request("ordem_fornecedor")
ordem_aval = request("ordem_aval")
ordem_especialidade = request("ordem_especialidade")
ordem_marca = request("ordem_marca")
ordem_periodo = request("ordem_periodo")
ordem_classificacao = request("ordem_classificacao")
'ordem_valor = request("ordem_valor")


sentido = request("sentido")
'if sentido = "DESC" Then
'ordem_mes = 0
'ordem_ano = 0
'Else
'end if

if campo_0 = 1 Then 'OR campo_2_2 = 1 Then
	codigo = "cd_codigo"
else
	codigo = "sequencia"
end if
%>
<!--#include file="../includes/ordem.asp"-->
<%
teste = request("teste")
response.write(teste)
ordem = "Order by sequencia " &primeiro&" "&segundo&" "&terceiro&" "&quarto&" "&quinto&" "&sexto&" "&setimo&" "&oitavo&" "&nono&" "&ultimo

mes_hoje = month(now)
ano_hoje = year(now)

mes_i = request("mes_i")
ano_i = request("ano_i")
dia_i = request("dia_i")
	if dia_i="" Then
		dia_i="1"
	end if
data_i = mes_i&"/"&dia_i&"/"&ano_i

mes_f = request("mes_f")
ano_f = request("ano_f")
dia_f = request("dia_f")

	ver_ultimo_dia = UltimoDiaMes(mes_f, ano_f) 
	if int(ver_ultimo_dia) < int(dia_f) then '*** verifica se o ultimo dia é válido.
	dia_f = UltimoDiaMes(mes_f, ano_f)
		alerta = ("O dia da data final selecionado foi substituído automaticamente pelo último dia válido do período.<br>")
	end if

	if dia_f = "" Then '*** Seleciona o ultimo dia do mes atual ao acessar o formulário vazio.
		dia_f = UltimoDiaMes(mes_hoje, ano_hoje)
		alerta = ("ultimo dia automático")
	end if	
	
data_f = mes_f&"/"&dia_f&"/"&ano_f


select_option = "num_qtd"

if cd_patrimonio_vazio = "1" Then
'cond_pat_vazio = " AND (cd_patrimonio IS not null)"
cond_pat_vazio = " AND (cd_patrimonio <> '')"
end if

if questao = "quantidade" then
	'selecao = "SELECT COUNT('"&select_option&"') as conta,cd_codigo ,"&campos&""
	selecao = "SELECT COUNT('"&select_option&"') as conta, "&codigo&", "&campos&""
	cond = " WHERE sequencia = '1' "
'	cond_ns = " AND (cd_ns <> ' ') AND (cd_ns is not null)"
	cond_pat = " AND (cd_patrimonio = '"&cd_patrimonio&"')"
	cond_patrimonio = " AND (cd_patrimonio = '"&nm_patrimonio&"')"
	'cond_pat_vazio = " AND (cd_patrimonio <> 0) "
	cond_pat_vazio = cond_pat_vazio
	cond_eqp = " AND (cd_equipamento = '"&cd_equipamento&"')"
	cond_und = " AND (nm_unidade = '"&nm_unidade&"')"
'	cond_mes = " AND (dt_mes BETWEEN '"&mes_i&"' AND '"&mes_f&"')"
'	cond_ano = " AND (dt_ano BETWEEN '"&ano_i&"' AND '"&ano_f&"')"
	cond_periodo = " AND (dt_os BETWEEN '"&data_i&"' AND '"&data_f&"')"
	cond_forn = " AND (cd_fornecedor = '"&cd_fornecedor&"')"
	cond_esp = " AND (cd_especialidade = '"&cd_especialidade&"')"
	cond_aval = " AND (cd_avaliacao = '"&cd_avaliacao&"')"
	cond_valor = " AND (cd_valor_orc = '"&cd_valor_orc&"')"
	cond_marca = " AND (cd_marca = '"&cd_marca&"')"
	agrupamento = " GROUP BY sequencia,"&campos&", "&codigo&""', cd_codigo"

Else
	selecao = "SELECT * FROM VIEW_os_lista"
End if

'***** Substitui a vírgula da sentença de busca *****
	busca_equip = Instr(1,nm_equipamento,",",0)
	if busca_equip <> 0 Then
		nm_equipamento = replace(nm_equipamento,",","%') OR (nm_equipamento Like '%")
	End if
'****************************************************

'If cd_ns = "0" Then
'	cond_ns = ""
'End if

If cd_equipamento = "0" AND nm_equipamento = "" Then
	cond_eqp = ""
ElseIf cd_equipamento = "0" AND nm_equipamento <> "" Then
	cond_eqp = "AND (nm_equipamento Like '%"&nm_equipamento&"%')"
End if

If campo_2_1 = "0" Then
	cond_cd = ""
End if

if nm_patrimonio = "" then 'if campo_2_3 = "0" Then
	cond_patrimonio = ""
Elseif nm_patrimonio <> "" Then
	'cond_patrimonio = "AND (cd_patrimonio Like '%"&nm_patrimonio&"%')"
	cond_patrimonio = "AND (cd_patrimonio = '"&nm_patrimonio&"')"
End if




If nm_unidade = "0" Then
	cond_und = ""
End if
'if ano_i = "0" OR ano_f = "0" Then
'	cond_ano_i = "AND (dt_ano >= '1950')"
'	cond_ano_f = "AND (dt_ano <= '3000')"
'end if
If cd_fornecedor = "0" then
	cond_forn = "" 
End if
If cd_especialidade = "0" then
	cond_esp = ""
End if
If cd_avaliacao = "0" then
	cond_aval = ""
End if
If cd_marca = "0" then
	cond_marca = ""
End if
%>
<body>
<br>
<table align="center" border="1">
	<tr>
		<td align="center" bgcolor="#C0C0C0"><b>RELATÓRIOS - <%=cd_patrimonio_vazio%><font color="red">Manutenção / Compras</font></b>
		</td>
	</tr>
		<tr><td align="center">		
			<table border="1" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" align="center">
				<form action="manutencao.asp" method="post" name="form">
				<input type="hidden" name="tipo" value="relatorio">
				<tr id="no_print">
					<td colspan="3"><img src="imagens/blackdot.gif" width="140" height="1"></td>
					<td colspan="4"><img src="imagens/blackdot.gif" width="485" height="1"></td>
					<td><img src="imagens/blackdot.gif" width="115" height="1"></td>
				</tr>
				<tr id="no_print">
					<td colspan="3">&nbsp;<b>MOSTRAR</b></td>
					<td colspan="4">&nbsp;<b>DETALHES</b></td>
					<td>&nbsp;<b>ORDEM</b></td>
				</tr>
				<tr id="no_print">
					<td colspan="3">&nbsp;<a href="javascript:selecionar_tudo()">tudo</a> :: <a href="javascript:deselecionar_tudo()">nada</a>
					</td>
					<td colspan="4">&nbsp;<b>QUANTIDADE</b>
						<input type="hidden" name="questao" value="quantidade">
						<!--select name="questao">
						<%'if questao="quantidade" Then%><%qtd = "Selected"%>
						<%'Elseif questao = "listagem" Then%><%list = "selected"%>
						<%'end if%>
							<option value="quantidade" <%=qtd%>>Quantidade</option>
							<option value="listagem" <%=list%>>Listagem</option>
						</select-->
					</td>
					<td><select name="ordem_conta" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_conta) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>
					</td>
					
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_0 <> 0 then%><%ck_campo_0 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="num_os" <%=ck_campo_0%>>Ordem de serviço
					</td>
					<td colspan="4"><input type="radio" name="os_pendente" value="0" <%=ck_pend1%>>&nbsp;Pendentes<br>
									<input type="radio" name="os_pendente" value="1" <%=ck_pend2%>>&nbsp;Encerradas<br>
									<input type="radio" name="os_pendente" value="" <%=ck_pend3%>>&nbsp;Todas
					</td>
					<td><select name="ordem_os" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_os) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
					
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_2_1 <> 0 then%><%ck_campo_2_1 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_ns" <%=ck_campo_2_1%>>Número de série
					</td>
					<td colspan="4">&nbsp;</td>		
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_ns" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_ns) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_2_2 <> 0 then%><%ck_campo_2_2 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_patrimonio" <%=ck_campo_2_2%>>Patrimônio
					</td>
					<%'***** Substitui o código de busca por vírgula *****
					'busca_patrimonio = Instr(1,nm_patrimonio,"%') OR (cd_patrimonio Like '%",0)
					'if busca_patrimonio <> 0 Then
					'cd_patrimonio = replace(nm_patrimonio,"%') OR (cd_patrimonio Like '%",",")
					'End if
					'****************************************************%>
					<%if cd_patrimonio_vazio = 1 then%><%patrimonio_vazio = "checked"%><%end if%>
					<td colspan="4"><input type="text" name="nm_patrimonio" value="" size="20">
					<input type="checkbox" name="cd_patrimonio_vazio" value="1" <%=patrimonio_vazio%>>(Não Vazios)  &nbsp;
						
						<%'select_1 = " top 100 percent * "
					'tabela = " View_patrimonio_lista"
					'condicoes = " Where dt_descarte Is Null  "
					'criterios = " Order by nm_equipamento"
					'strsql ="up_listagem @selecao='"&select_1&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"     '"TBL_equipamento"
										  	
					  	'Set rs_pat = dbconn.execute(strsql)%>
						<!--select name="cd_patrimonio_0" class="inputs"><option value="">Selecione</option>
						<%'Do while Not rs_pat.EOF %>
						<%'patrimonio=rs_pat("cd_patrimonio")%>
						<%'if nm_pat=patrimonio Then%>
						<%'ck_pat="selected"%>
						<%'else ck_pat=""%>
						<%'end if%>
						<option value="<%'=rs_pat("cd_codigo")%>" <%'=ck_pat%>><%'=rs_pat("cd_patrimonio")%></option><%'rs_pat.movenext
					  	'Loop%>-->
						
					</td>		
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_patrimonio" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_patrimonio) = numero Then
							check = "selected"
						end if%>						
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_1 <> 0 then%><%ck_campo_1 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_equipamento,nm_equipamento" <%=ck_campo_1%>>Equipamentos
					</td>
					<td colspan="4">
						<select name="cd_equipamento">
							<option value="0">Todos
						<%xsql_eqp = "Select * From TBL_equipamento order by nm_equipamento"
						Set rs_eqp = dbconn.execute(xsql_eqp)
						Do while not rs_eqp.EOF
							cd_equip = rs_eqp("cd_codigo")
							nm_equip = rs_eqp("nm_equipamento")
								if int(cd_equipamento) = cd_equip then
									selected = "selected" 
								End if%>
							<option value="<%=cd_equip%>" <%=selected%>><%=nm_equip%>
							</option>
						<%rs_eqp.movenext
						selected = ""
						loop%>
						</select>
		<%'***** Substitui o código de busca por vírgula *****
		busca_equip = Instr(1,nm_equipamento,"%') OR (nm_equipamento Like '%",0)
		if busca_equip <> 0 Then
		nm_equipamento = replace(nm_equipamento,"%') OR (nm_equipamento Like '%",",")
		End if
		'****************************************************%>
						 ou <input type="text" name="nm_equipamento" value="<%=nm_equipamento%>" size="20">
					</td>
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_equip" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_equip) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_2 <> 0 then%><%ck_campo_2 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="nm_unidade" <%=ck_campo_2%>>Unidades
					</td>
					<td colspan="4">
						<select name="nm_unidade">
							<option value="0">Todos
						<%xsql_und = "Select * From TBL_unidades where cd_status = 1 order by nm_unidade"
						Set rs_und = dbconn.execute(xsql_und)
						Do while not rs_und.EOF
							nm_sigla = rs_und("nm_sigla")
							nm_unid = rs_und("nm_unidade")
								if nm_unidade = nm_sigla then
									selected = "selected" 
								End if%>
							<option value="<%=nm_sigla%>" <%=selected%>><%=nm_unid%> - <%=nm_sigla%>
							</option>
						<%rs_und.movenext
						selected = ""
						loop
						
						nm_unidade = ""%>
						</select>
					</td>
					<td><!--input type="text" name="ordem_unidade" value="" size="1"-->
						<select name="ordem_unidade" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_unidade) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_3 <> 0 then%><%ck_campo_3 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_fornecedor,nm_fornecedor" <%=ck_campo_3%>>A. Técnica
					</td>
					<td colspan="4">
						<select name="cd_fornecedor">
							<option value="0">Todos
						<%xsql_forn = "Select * From TBL_fornecedor order by nm_fornecedor"
						Set rs_forn = dbconn.execute(xsql_forn)
						Do while not rs_forn.EOF
							cd_forn = rs_forn("cd_codigo")
							nm_forn = rs_forn("nm_fornecedor")
								if int(cd_fornecedor) = cd_forn then
									selected = "selected" 
								End if%>
							<option value="<%=cd_forn%>" <%=selected%>><%=nm_forn%>
							</option>
						<%rs_forn.movenext
						selected = ""
						loop%>
						</select>
					</td>
					<td><!--input type="text" name="ordem_fornecedor" value="" size="1"-->
					<select name="ordem_fornecedor" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_fornecedor) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_4 <> 0 then%><%ck_campo_4 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_especialidade,nm_especialidade" <%=ck_campo_4%>>Especialidades
					</td>
					<td colspan="4">
						<select name="cd_especialidade">
							<option value="0">Todos
						<%xsql_esp = "Select * From TBL_especialidades order by nm_especialidade"
						Set rs_esp = dbconn.execute(xsql_esp)
						Do while not rs_esp.EOF
							cd_esp = rs_esp("cd_codigo")
							nm_esp = rs_esp("nm_especialidade")
								if int(cd_especialidade) = cd_esp then
									selected = "selected" 
								End if
							%>
							<option value="<%=cd_esp%>" <%=selected%>><%=nm_esp%> <%'=cd_esp%>
							</option>
						<%rs_esp.movenext
						selected = ""
						loop%>
						</select>
						</td>
						<td>
						<select name="ordem_especialidade" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_especialidade) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>
						</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_5 <> 0 then%><%ck_campo_5 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_avaliacao,nm_avaliacao" <%=ck_campo_5%>>Avaliação
					</td>
					<td colspan="4">
						<select name="cd_avaliacao">
							<option value="0">Todos
						<%xsql_aval = "Select * From TBL_avaliacao Order by nm_avaliacao"
						Set rs_aval = dbconn.execute(xsql_aval)
						Do while not rs_aval.EOF
							cd_aval = rs_aval("cd_codigo")
							nm_aval = rs_aval("nm_avaliacao")
								if int(cd_avaliacao) = cd_aval then
									selected = "selected" 
								End if
								%>
							<option value="<%=cd_aval%>" <%=selected%>><%=nm_aval%>
							</option>
						<%rs_aval.movenext
						selected = ""
						loop%>
						</select>
					</td>
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_aval" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_aval) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
				<tr id="no_print">
					<td colspan="3"><%if campo_6 <> 0 then%><%ck_campo_6 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_marca,nm_marca" <%=ck_campo_6%>>Marcas
					</td>
					<td colspan="4">
						<select name="cd_marca">
							<option value="0">Todos
						<%xsql_marca = "Select * From TBL_marca Order by nm_marca"
						Set rs_marca = dbconn.execute(xsql_marca)
						Do while not rs_marca.EOF
							cdmarca = rs_marca("cd_codigo")
							marca = rs_marca("nm_marca")
								if int(cd_marca) = cdmarca then
									selected = "selected" 
								End if%>
							<option value="<%=cdmarca%>" <%=selected%>><%=marca%>
							</option>
						<%rs_marca.movenext
						selected = ""
						loop%>
						</select>
					</td>
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_marca" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_marca) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
				
				<tr id="no_print">
					<td colspan="3"><%if campo_8 <> 0 then%><%ck_campo_8 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="num_os_assist" <%=ck_campo_8%>>Classificação
					</td>
					<td colspan="4">&nbsp;</td>
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_classificacao" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_classificacao) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- --------------------------------------------------------------------------------- -->				
				<tr id="no_print">
					<td colspan="3"><%if campo_12 <> 0 then%><%ck_campo_12 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="cd_valor_orc" <%=ck_campo_12%>>Valores
					</td>
					<td colspan="4">&nbsp;</td>
					<td><!--input type="text" name="ordem_equip" value="" size="1"-->
					<select name="ordem_valor_orc" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_valor) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- --------------------------------------------------------------------------------- -->				



				<tr id="no_print">
					<td colspan="3"><%if campo_7 <> 0 then%><%ck_campo_7 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value="dt_os" <%=ck_campo_7%>>Período 
					</td>
					<td colspan="4">
						<select name="dia_i">
							<%numero = 1
							do while NOT numero = 32
							if int(dia_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_i">
							<%numero = 1
							do while NOT numero = 13
							if int(mes_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<%if ano_i = "" then%><%ano_i=1900%><%end if%>
						<input type="text" name="ano_i" size="4" maxlength="4" value="<%=ano_i%>">
						até
						<select name="dia_f">
							<%numero = 1
							do while NOT numero = 32
								'if dia_f = "" then
								'teste = "nada"
								'Else
								'teste = "algo"
								'end if
								
								if int(dia_f) = numero Then
								check = "selected"
								'elseif
								end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_f">
							<%numero = 1
							do while NOT numero = 13
								if mes_f = "" Then
									if mes_hoje = numero Then
									'numero = 12
									check = "selected"
									end if
								Elseif int(mes_f) = numero Then
									check = "selected"
								end if%>		
							<option value="<%=numero%>" <%=check%>><%=left(mes_selecionado(numero),3)%></option>
							<%numero = numero+1
						check = ""
						loop%>							
						</select> /
						<%if ano_f = "" then%><%ano_f=ano_atual%><%end if%> 
						<input type="text" name="ano_f" size="4" maxlength="4" value="<%=ano_f%>"> 
					</td>
					<td><select name="ordem_periodo" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_periodo) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%>º</option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>				
				<tr id="no_print">
					<td colspan="7">&nbsp;<font color="#ff0000"><%=alerta%> <%'=dia_f%></font></td>
					<td>
					<select name="sentido" size="1"><%if sentido = "DESC" Then%><%desc="selected"%><%end if%>
						<option value="ASC">Crescente</option>
						<option value="DESC" <%=desc%>>Decrescente</option>
					</select>
					</td>
				</tr>		
				<tr id="no_print">
					<td colspan="8" align="center">
						<input type="hidden" name="x" value="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="submit" value="Gerar Relatório">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input onclick="javascript:JsRedefinir('')" type="button" value="Limpar" title="Redefinir"> 
					</td>
				</tr>
			</table>
			<blockquote>	
			<table width="2" align="center" border="0" >
				<!--tr>
					<td colspan="100%">&nbsp;
					</td>
				</tr-->
			</form>
			<%
			'******************************
			'*							  *
			'* 	  1ª Parte - formulário   *
			'*                            *
			'******************************
		If campos = "" Then%>
				<tr>
					<td colspan="100%"><br><%response.write("<b> Nenhum item foi selecionado</b>")%>
					</td>
				</tr>
		<%Else%>
				<tr bgcolor="#c0c0c0">
					<td>&nbsp;Qtd</td>
<%if campo_0 <> 0 Then%><td>&nbsp;O.S.</td><%end if%>
<%if campo_7 <> 0 Then%><td>&nbsp;Mes</td><%end if%>
<%if campo_1 <> 0 Then%><td>&nbsp;Equipamento</td><%end if%>
<%if campo_2_1 <> 0 Then%><td>&nbsp;N/S</td><%end if%>
<%if campo_2_2 <> 0 Then%><td>&nbsp;Patrimônio</td><%end if%>
<%if campo_2 <> 0 Then%><td>&nbsp;Unidade</td><%end if%>
<%if campo_3 <> 0 Then%><td>&nbsp;A. Tecnica</td><%end if%>
<%if campo_4 <> 0 Then%><td>&nbsp;Especialidades</td><%end if%>
<%if campo_5 <> 0 Then%><td>&nbsp;Avaliacao</td><%end if%>
<%if campo_6 <> 0 Then%><td>&nbsp;Marca</td><%end if%>
<%if campo_8 <> 0 Then%><td>&nbsp;Class.</td><%end if%>
<%if campo_12 <> 0 Then%><td>&nbsp;Valor</td><%end if%>
				</tr>
				<tr><td colspan="100%" id="ok_print"><img src="imagens/blackdot.gif" alt="" width="100%" height="1"></td></tr>
			<%
			
			'xsql = ""&selecao&" FROM VIEW_os_lista "&cond&" "&cond_mes&" "&cond_ano&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&agrupamento&" "&ordem&" "&sentido&""'
			xsql = ""&selecao&" FROM VIEW_os_lista "&cond&" "&cond_periodo&" "&cond_ns&" "&cond_patrimonio&" "&cond_pat_vazio&" "&cond_eqp&" "&cond_und&" "&cond_forn&" "&cond_esp&" "&cond_aval&" "&cond_marca&" "&cond_fecha&" "&agrupamento&" "&ordem&" "&sentido&""'
			Set rs = dbconn.execute(xsql)
			do while Not rs.EOF 
				if campo_0 = 1 then
					cd_cod = rs("cd_codigo")
				end if
			
			If questao = "quantidade" then
			conta = rs("conta")
			
			Else
			
			'num_qtd = rs("num_qtd")
			'nm_unidade = rs("nm_unidade")
			'nm_fornecedor = rs("nm_fornecedor")
			'nm_marca = rs("nm_marca")
			'nm_especialidade = rs("nm_especialidade")
			'cd_fornecedor = rs("cd_fornecedor")
			'nm_avaliacao = rs("nm_avaliacao")
			'nm_marca = rs("nm_marca")
			
			end if
			
				'fecha = rs("fecha")
			
			if campo_0 <> 0 Then
				num_os = rs("num_os")
			End if
			if campo_2_1 <> 0 Then
				cd_ns = rs("cd_ns")
			End if
			if campo_2_2 <> 0 Then
				cd_patrimonio = rs("cd_patrimonio")
			End if
			if campo_1 <> 0 Then
				nm_equipamento = rs("nm_equipamento")
			End if
			if campo_2 <> 0 Then
				nm_unidade = rs("nm_unidade")
			end if
			if campo_3 <> 0 Then
				nm_fornecedor = rs("nm_fornecedor")
			end if
			if campo_4 <> 0 Then
				nm_especialidade = rs("nm_especialidade")
			end if
			if campo_5 <> 0 Then
				nm_avaliacao = rs("nm_avaliacao")
			end if
			if campo_6 <> 0 Then
				nm_marca = rs("nm_marca")
			end if
			
			if campo_7 <> 0 Then
				dt_os = rs("dt_os")
			end if
			
			if campo_8 <> 0 Then
				num_os_assist = rs("num_os_assist")
			end if
			
			dt_os_d = zero(day(dt_os))
			dt_os_m = zero(month(dt_os))
			dt_os_a = year(dt_os)
			
			dt_os = dt_os_d&"/"&dt_os_m&"/"&dt_os_a%>
			
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:;" onmouseout="mOut(this,'');">
					<td valign="top">&nbsp;<%=zero(conta)%></td>
<%if campo_0 <> 0 Then%><td valign="top"><a href="javascript:;" onclick="window.open('manutencao/manutencao_andamento.asp?cd_codigo=<%=cd_cod%>&campo=cd_codigo&visual=1&jan=1', 'principal', 'width=600, height=500,scrollbars=1'); return false;"><%=num_os%></a></td><%end if%>
<%if campo_7 <> 0 Then%><td valign="top"><%=dt_os%><%=left(mes_selecionado(dt_mes),3)%></td><%end if%>
<%if campo_1 <> 0 Then%><td valign="top"><%=nm_equipamento%></td><%end if%>
<%if campo_2_1 <> 0 Then%><td valign="top">&nbsp;<%=cd_ns%></td><%end if%>
<%if campo_2_2 <> 0 Then%><td valign="top">&nbsp;<%=cd_patrimonio%></td><%end if%>
<%if campo_2 <> 0 Then%><td valign="top">&nbsp;<%=nm_unidade%></td><%end if%>
<%if campo_3 <> 0 Then%><td valign="top"><%=nm_fornecedor%></td><%end if%>
<%if campo_4 <> 0 Then%><td valign="top">&nbsp;<%=nm_especialidade%></td><%end if%>
<%if campo_5 <> 0 Then%><td valign="top">&nbsp;<%=nm_avaliacao%></td><%end if%>
<%if campo_6 <> 0 Then%><td valign="top"><%=nm_marca%></td><%end if%>
<%if campo_8 <> 0 Then%><td valign="top"><%=num_os_assist%></td><%end if%>
<td valign="top"><%=fecha%></td>
				</tr>
			<%rs.movenext
			geral = geral + conta
			loop
			
			'end if%>
			
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>					
			<tr bgcolor="#C0C0C0">
					<td colspan="100%">&nbsp;Total: <%=geral%></td>
				</tr>
		<%End if%>
				<tr id="ok_print">
						<td><img src="imagens/px.gif" width="25" height="1"></td>
<%if campo_0 <> 0 Then%><td><img src="imagens/px.gif" width="30" height="1"></td><%end if%>
<%if campo_7 <> 0 Then%><td><img src="imagens/px.gif" width="65" height="1"></td><%end if%>
<%if campo_1 <> 0 Then%><td><img src="imagens/px.gif" width="140" height="1"></td><%end if%>
<%if campo_2_1 <> 0 Then%><td><img src="imagens/px.gif" width="75" height="1"></td><%end if%>
<%if campo_2_2 <> 0 Then%><td><img src="imagens/px.gif" width="75" height="1"></td><%end if%>
<%if campo_2 <> 0 Then%><td><img src="imagens/px.gif" width="50" height="1"></td><%end if%>
<%if campo_3 <> 0 Then%><td><img src="imagens/px.gif" width="80" height="1"></td><%end if%>
<%if campo_4 <> 0 Then%><td><img src="imagens/px.gif" width="80" height="1"></td><%end if%>
<%if campo_5 <> 0 Then%><td><img src="imagens/px.gif" width="65" height="1"></td><%end if%>
<%if campo_6 <> 0 Then%><td><img src="imagens/px.gif" width="50" height="1"></td><%end if%>
				</tr>
			<tr id="no_print">
						<td><img src="imagens/px.gif" width="25" height="1"></td>
<%if campo_0 <> 0 Then%><td><img src="imagens/px.gif" width="40" height="1"></td><%end if%>
<%if campo_7 <> 0 Then%><td><img src="imagens/px.gif" width="70" height="1"></td><%end if%>
<%if campo_1 <> 0 Then%><td><img src="imagens/px.gif" width="160" height="1"></td><%end if%>
<%if campo_2_1 <> 0 Then%><td><img src="imagens/px.gif" width="90" height="1"></td><%end if%>
<%if campo_2_2 <> 0 Then%><td><img src="imagens/px.gif" width="90" height="1"></td><%end if%>
<%if campo_2 <> 0 Then%><td><img src="imagens/px.gif" width="60" height="1"></td><%end if%>
<%if campo_3 <> 0 Then%><td><img src="imagens/px.gif" width="100" height="1"></td><%end if%>
<%if campo_4 <> 0 Then%><td><img src="imagens/px.gif" width="95" height="1"></td><%end if%>
<%if campo_5 <> 0 Then%><td><img src="imagens/px.gif" width="80" height="1"></td><%end if%>
<%if campo_6 <> 0 Then%><td><img src="imagens/px.gif" width="65" height="1"></td><%end if%>
				</tr>
				<%if session_cd_usuario = 46 Then%>
				<tr class="txt">
		<td align="left" colspan="100%">< Tabela principal><br>(<%=campo_2_2%>)
		<b>SELECT COUNT</b>(<%=select_option%>) as conta,<%=campos%><br>
		<b>FROM VIEW_os_lista </b><%=cond%> <%=cond_periodo%> <%'=cond_ano%> <%=cond_ns%> <%=cond_eqp%> <%'=cond_pat%> <%=cond_patrimonio%> <%=cond_und%> <%=cond_forn%> <%=cond_esp%> <%=cond_aval%> <%=cond_marca%><br>
		<%=agrupamento%><br>
		 o:<%=ordem%> s:<%=sentido%><br><%="++ "&ordem_patrimonio&" ++"%><br>
		 (1-<%=primeiro%>), (2-<%=segundo%>), (3-<%=terceiro%>), (4-<%=quarto%>), (5-<%=quinto%>), (6-<%=sexto%>), (7-<%=setimo%>), (8<%=oitavo%>), (9-<%=nono%>), 10-(<%=decimo%>), (U- ""%>)
		<br>>
		</td>
	</tr><%end if%>
				<!--tr><td colspan="100%">&nbsp;(<%=campos%>)</td></tr-->
			</table>
		</td><!-- Fim da tabela principal-->
	</tr>
</table></blockquote><br>
<br>

<SCRIPT LANGUAGE="javascript">

function JsRedefinir(cod_a, cod_b, cod_c, cod_d)
{
		document.location.href('relatorio_manut.asp?acao=redefinir');
}

</SCRIPT>




</body>
