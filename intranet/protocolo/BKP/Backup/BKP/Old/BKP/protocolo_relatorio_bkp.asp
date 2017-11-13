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
	width: 690;
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
// -->	
</script>
<%
'**********************************************
'**************  Variáveis  *******************
'**********************************************
campos = request("campos")
ok = request("ok")


'******************** Unidade ********************
cd_unidade = request("cd_unidade")
nm_unidade = request("nm_unidade")
	'Verifica as unidades solicitadas e monta o comando para o SQL
		'campos = campos&", nm_unidade, nm_sigla, a053_codung "
		numero = 1
		proc_array = split(nm_unidade,",")
		For Each proc_item In proc_array
				if numero = 1 Then
				nome_unidade = proc_item
				elseif numero >= 2 then
				nome_unidade = nome_unidade&"%' or nm_unidade like '%"&proc_item
				end if
		numero = numero + 1
		Next
		
	if nm_unidade <> ""  Then
		nome_unidade = " AND nm_unidade like '%"&nome_unidade&"%'"
		campos = campos&" nm_unidade, nm_sigla, a053_codung "
		cd_unidade = "0"
	Elseif cd_unidade > "0" then
		unidade = "AND a053_codung = "&cd_unidade&""
		campos = campos&",a053_codung, nm_sigla, nm_unidade "
	end if

'******************** Idade ********************
a002_pacida_i = request("a002_pacida_i")
a002_pacida_f = request("a002_pacida_f")
	if a002_pacida_i <> "" AND a002_pacida_f <> "" Then
		idade = "AND a002_pacida BETWEEN "&a002_pacida_i&" AND "&a002_pacida_f&""
	Elseif a002_pacida_i <> "" AND a002_pacida_f = "" Then
		idade = "AND a002_pacida BETWEEN "&a002_pacida_i&" AND "&a002_pacida_i&""
	end if

'******************** Sexo ********************
cd_sexo = request("cd_sexo")
	if cd_sexo <> "" then
		sexo = "AND a002_pacsex='"&cd_sexo&"'"
	end if

'******************** Convenio ********************
cd_convenio = request("cd_convenio")
	if cd_convenio <> "0" Then
		convenio = "AND a054_codcon = "&cd_convenio&""
	end if

'******************** Medico ********************
'cd_medico = request("cd_medico")

'**********************************************
'** Verifica quais campos foram selecionados **
'**********************************************
	
	'campo_0 = Instr(1,campos,"conta",0)
	campo_2 = Instr(1,campos,"a053_codung",0)
		campo_2_1 = Instr(1,campos,"nm_unidade",0)
		campo_2_2 = Instr(1,campos,"sigla",0)
	campo_3 = Instr(1,campos,"a002_numpro",0)
	campo_4 = Instr(1,campos,"a002_pacreg",0)
	campo_5 = Instr(1,campos,"a002_pacida",0)
	campo_6 = Instr(1,campos,"a002_pacsex",0)
	campo_7 = Instr(1,campos,"a054_codcon",0)
	
	campo_9_1 = Instr(1,campos,"dt_mes",0)
	campo_9_2 = Instr(1,campos,"a002_datpro",0)

'********************************************
'************** Período ***********************
dt_dia = request("dt_dia")
'response.write("dt_dia")

dia_atual = day(now)
mes_atual = month(now)
ano_atual = year(now)

dia_i = request("dia_i")
mes_i = request("mes_i")
ano_i = request("ano_i")
		'***
dia_f = request("dia_f")
mes_f = request("mes_f")
ano_f = request("ano_f")

	'ultimo_dia = ultimodiames(mes_f,ano_f)
	if ultimodiames(mes_f,ano_f) < int(dia_f) then
		dia_f = ultimodiames(mes_f,ano_f)
		'response.write("muda data")
	end if
		'***
		data_i = mes_i&"/"&dia_i&"/"&ano_i
		data_f = mes_f&"/"&dia_f&"/"&ano_f
			'***
			
		if campo_9_1 > 0 then 'Mensal
			If ano_i&zero(mes_i) <= ano_f&zero(mes_f) Then '1 - Menor ou igual
			'periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_f&"'"
			periodo = "(dt_ano BETWEEN '"&ano_i&"' AND '"&ano_f&"') AND (dt_mes BETWEEN '"&mes_i&"' AND '"&mes_f&"')"
			elseIf ano_i&zero(mes_i)&zero(dia_i) > ano_f&zero(mes_f)&zero(dia_f) Then '1 - Maior
			periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_i&"'"
			ok = ""
			periodo_erro = "O período selecionado é inválido!<br>A data inicial é maior que a final."
			end if
			'response.write("mensal")
		Elseif campo_9_2 > 0 then 'diario
			If ano_i&zero(mes_i)&zero(dia_i) <= ano_f&zero(mes_f)&zero(dia_f) Then '1 - Menor ou igual
			periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_f&"'"
			elseIf ano_i&zero(mes_i)&zero(dia_i) > ano_f&zero(mes_f)&zero(dia_f) Then '1 - Maior
			periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_i&"'"
			ok = ""
			periodo_erro = "O período selecionado é inválido!<br>A data inicial é maior que a final."
			end if
			'response.write("diário")
		End if
	
'**********************************************
'************** Seleção ***********************
'**********************************************
selecao = " COUNT (a002_numpro) AS conta"


if campos = "" Then
selecao = selecao
Else
selecao = selecao &", "&campos 
End if
	'Elimina uma possível duplicação de virgulas
	selecao = replace(selecao,", ,",",")

'**********************************************
'************** Condições *********************
condicao = "Where "&periodo&" "&unidade&" "&nome_unidade&" "&idade&" "&sexo&" "&convenio&""

'**********************************************
'************** Agrupamento *********************
if campos <> "" Then
agrupamento = "Group by "&campos
	teste_agrupamento = Instr(1,agrupamento,"Group by ,",0)
			if teste_agrupamento = 1 then
			agrupamento = replace(agrupamento,"Group by ,","Group by ")'Retira a virgula
			end if
			'response.write(teste_agrupamento)

'**********************************************
'**************** Ordem ***********************
ordem_qtd = request("ordem_qtd")
ordem_unidade = request("ordem_unidade")
ordem_protocolo = request("ordem_protocolo")
ordem_paciente = request("ordem_paciente")
ordem_idade = request("ordem_idade")
ordem_sexo = request("ordem_sexo")
ordem_convenio = request("ordem_convenio")
ordem_data = request("ordem_data")
%>

<!--#include file="../includes/ordem_protocolos.asp"-->
	<%if ordem <> "" Then
		ordem = "ORDER BY "&ordem
			teste_ordem = Instr(1,ordem,"ORDER BY ,",0)
			if teste_ordem = 1 then
			ordem = replace(ordem,",","")'Retira a virgula
			end if 
	  end if

end if%>

<body>
<br>
<table align="center" border="1">
	<tr>
		<td align="center" bgcolor="#C0C0C0"><b>RELATÓRIOS - <%=cd_patrimonio_vazio%><font color="red">Protocolos</font></b>
		</td>
	</tr>
		<tr><td align="center">		
			<table border="1" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" align="center">
				<form action="protocolo.asp" method="post" name="form">
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
<!-- *** Quantidade *** -->
				<tr id="no_print">
					<td colspan="3">&nbsp; &nbsp; &nbsp; Quantidade </td>
					<td colspan="4">&nbsp;
					</td>
					<td><select name="ordem_qtd" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_qtd) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Unidade *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_2 <> 0 then%><%ck_campo_2 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="a053_codung, nm_sigla" <%=ck_campo_2%>>&nbsp;Unidade 
					</td>
					<td colspan="4">
					<select name="cd_unidade" class="inputs">
						<option value="0"></option>
							<%strsql = "SELECT * FROM TBL_unidades WHERE cd_codigo <> '6' AND cd_status > '0' Order by 'nm_unidade'"
							Set rs_uni = dbconn.execute(strsql)
							
							while not rs_uni.EOF
							cd_uni = rs_uni("cd_codigo")
							nm_sigla = rs_uni("nm_sigla")
							nome_unidade = rs_uni("nm_unidade")
							
							if int(cd_unidade) = cd_uni Then
							ck_uni = "selected"
							end if%>						
						<option value="<%=cd_uni%>" <%=ck_uni%>><%=nm_sigla%> - <%=nome_unidade%></option>
							<%rs_uni.movenext
							ck_uni = ""
							wend
							%>
					</select>
					OU
					<input type="text" name="nm_unidade" value="<%=nm_unidade%>" size="20" maxlength="200"  class="inputs">
					</td>
					<td><select name="ordem_unidade" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_unidade) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Nº Protocolo *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_3 <> 0 then%><%ck_campo_3 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="a002_numpro,a053_codung" <%=ck_campo_3%>>&nbsp;Protocolo						
					</td>
					<td colspan="4"><input type="text" name="a002_numpro" size="70" maxlength="500">
						
					</td>
					<td><select name="ordem_protocolo" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_protocolo) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Nº paciente *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_4 <> 0 then%><%ck_campo_4 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="a002_pacreg, a002_pacnom" <%=ck_campo_4%>>&nbsp;Paciente						
					</td>
					<td colspan="4"><input type="text" name="a002_pacnom" size="70" maxlength="500">
						
					</td>
					<td><select name="ordem_paciente" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_paciente) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Idade *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_5 <> 0 then%><%ck_campo_5 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="a002_pacida" <%=ck_campo_5%>>&nbsp;Idade
					</td>
					<td colspan="4">				
						<input type="text" size="3" name="a002_pacida_i" maxlength="3" value="<%=a002_pacida_i%>"> até 
						<input type="text" size="3" name="a002_pacida_f" maxlength="3" value="<%=a002_pacida_f%>">
					</td>
					<td><select name="ordem_idade" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_idade) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Sexo *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_6 <> 0 then%><%ck_campo_6 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="a002_pacsex" <%=ck_campo_6%>>&nbsp;Sexo
					</td>
					<td colspan="4"><%if cd_sexo = "F" Then%><%ck_fem = "selected"%><%Elseif cd_sexo = "M" Then%><%ck_masc = "selected"%><%end if%>
					<select name="cd_sexo" size="1">
						<option value=""></option>
						<option value="F" <%=ck_fem%>>F</option>
						<option value="M" <%=ck_masc%>>M</option>
					</select>
					</td>
					<td><select name="ordem_sexo" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_sexo) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Convenio *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_7 <> 0 then%><%ck_campo_7 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="a054_codcon" <%=ck_campo_7%>>&nbsp;Convênio 
					</td>
					<td colspan="4">
					<select name="cd_convenio" class="inputs">
						<option value="0"></option>
							<%strsql = "SELECT * FROM TBL_convenio order by nm_convenio"
							Set rs_conv = dbconn.execute(strsql)
							
							while not rs_conv.EOF
							cd_conv = rs_conv("cd_codigo")
							nm_convenio = rs_conv("nm_convenio")
							
							if int(cd_convenio) = cd_conv Then
							ck_conv = "selected"
							end if%>						
						<option value="<%=cd_conv%>" <%=ck_conv%>><%=nm_convenio%></option>
							<%rs_conv.movenext
							ck_conv = ""
							wend
							%>
					</select>
					</td>
					<td><select name="ordem_convenio" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_convenio) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>
<!-- *** Data *** -->
				<tr id="no_print">
					<td colspan="3">
						<%if campo_9_1 => 0 then%><%ck_campo_9_1 ="checked"%><%else%><%ck_campo_9_2 ="checked"%><%end if%>
						<input type="radio" name="campos" value="dt_ano, dt_mes" <%=ck_campo_9_1%>>mensal
						<input type="radio" name="campos" value="a002_datpro" <%=ck_campo_9_2%>>diário
						
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
						<%if ano_i = "" then%><%ano_i=ano_atual%><%end if%>
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
					<td><select name="ordem_data" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 12
						if int(ordem_data) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
				</tr>

				
<!-- ***********************************************************************************************-->
<!-- *********************************      ORDEM      *********************************************-->
<!-- ***********************************************************************************************-->

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
						<input type="submit" name="ok" value="Gerar Relatório">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input onclick="javascript:JsRedefinir('')" type="button" value="Limpar" title="Redefinir"> 
					</td>
				</tr>
				<tr class="txt">
					<td align="left" colspan="100%">
					<font color="red"><%=periodo_erro%></font><br>
					</td>
				</tr>
				<tr>
					<td align="left" colspan="100%" bgcolor="#ffff00">
					<b>SELECT </b><%=selecao%><br>
					<b>FROM VIEW_protocolo_lista </b><br>
					<%=condicao%> <%=cond_periodo%> <%=cond_und%> <br>
					<%=agrupamento%><br>
					<%=ordem%><br>
					</td>
				</tr>
			</table>
			<br>
			<table width="2" align="center" border="0" >
				<!--tr>
					<td colspan="100%">&nbsp;
					</td>
				</tr-->
			</form>
	<%if ok <> "" Then 'Inicio do resultado da pesquisa
			'******************************
			'*							  *
			'* 	  2ª Parte - Resultados   *
			'*                            *
			'******************************	%>
				<tr bgcolor="#c0c0c0">
					<td>&nbsp;Qtd</td>
					<td>&nbsp;Unid.</td>
					<td>&nbsp;Protocolo</td>
					<td colspan="2">&nbsp;Paciente</td>
					<td>&nbsp;Idade</td>					
					<td>&nbsp;Sexo</td>
					<td>&nbsp;Convênio</td>
					<td>&nbsp;Data</td>
				</tr>
				<tr><td colspan="100%" id="ok_print"><img src="imagens/blackdot.gif" alt="" width="100%" height="1"></td></tr>
			<%xsql = "SELECT "&selecao&" FROM VIEW_protocolo_lista "&condicao&" "&agrupamento&" "&ordem&" "&sentido&" "						
			Set rs = dbconn.execute(xsql)
			
			do while Not rs.EOF 
			
			conta = rs("conta")
			
			if campo_2 <> 0 then ' Unidade
				a053_codung = rs("a053_codung")
				
			end if
				if campo_2_1 <> 0 Then ' Nome da Unidade
				nm_unidade = rs("nm_sigla")
				end if
			
			if campo_3 <> 0 then ' Nº protocolo
				a002_numpro = rs("a002_numpro")
				a053_codung = rs("a053_codung")
			end if
			
			if campo_4 <> 0 then ' Cod. Paciente
				a002_pacreg = rs("a002_pacreg")
				a002_pacnom = rs("a002_pacnom")
			end if
			
			if campo_5 <> 0 then ' Idade
				a002_pacida = rs("a002_pacida")
			end if
			
			if campo_6 <> 0 then ' Sexo
				a002_pacsex = rs("a002_pacsex")
			end if
			
			if campo_7 <> 0 then ' Convenio
				a054_codcon = rs("a054_codcon")
			end if
			
			if campo_9_1 <> 0 then
			
			dt_mes = rs("dt_mes")
				dt_ano = rs("dt_ano")
				A002_datpro = zero(dt_mes)&"/"&dt_ano
			Else
				A002_datpro_dia = day(rs("A002_datpro"))
				A002_datpro_mes = month(rs("A002_datpro"))
				A002_datpro_ano = year(rs("A002_datpro"))
					A002_datpro = zero(A002_datpro_dia)&"/"&zero(A002_datpro_mes)&"/"&A002_datpro_ano
			end if
			%>
			
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:;" onmouseout="mOut(this,'');">
					<td valign="top">&nbsp;<%=zero(conta)%></td>
					<td valign="top">&nbsp;<%'=zero(A053_codung)%><%=nm_unidade%></td>
					<td valign="top">&nbsp;<%=zero(A053_codung)%>.<%=zero(a002_numpro)%></td>
					<td valign="top">&nbsp;<%=a002_pacreg%></td><td><%=a002_pacnom%></td>
					<td valign="top">&nbsp;<%=a002_pacida%></td>
					<td valign="top">&nbsp;<%=a002_pacsex%></td>
					<td valign="top">&nbsp;<%=a054_codcon%></td>
					<td valign="top">&nbsp;<%=A002_datpro%></td>
				</tr>
			<%rs.movenext
			geral = geral + conta
			loop
			
			'end if%>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>					
			<tr bgcolor="#C0C0C0">
					<td colspan="100%">&nbsp;Total: <%=geral%></td>
				</tr>
		<%'End if%>
				
		
	<%'end if 'Fim do resultado da pesquisa%>	
		
		<tr bgcolor="#C0C0C0">
		<td align="left" colspan="100%">
		<!--
		<b>SELECT </b><%=selecao%><br>
		<b>FROM VIEW_protocolo_lista </b><br>
		<%=condicao%> <%=cond_periodo%> <%=cond_und%> <br>
		<%=agrupamento%><br>
		<%=ordem%><br>
		<br>
		<br>-->
		 Campos: <%=campos%><br>
		 o:<%=ordem%><br> 
		 s:<%=sentido%><br><br>
		 (1-<%=primeiro%>), (2-<%=segundo%>), (3-<%=terceiro%>), (4-<%=quarto%>), (5-<%=quinto%>), (6-<%=sexto%>), (7-<%=setimo%>), (8<%=oitavo%>), (9-<%=nono%>), 10-(<%=decimo%>), (U-<%'""%>)
		<br>
		campos: (<%=campo_3%>) * <br>
		unidade(<%=campo_2_2%>)<br>
		dia_f: (<%=dia_f%>)<br>
		Ok: <%=ok%>
		</td>
	</tr><%'end if%>
	<%end if 'Fim do resultado da pesquisa%>
				<tr><td colspan="100%">&nbsp;</td></tr>
			</table>
		</td><!-- Fim da tabela principal-->
	</tr>
</table><br>
<br>

<SCRIPT LANGUAGE="javascript">

function JsRedefinir(cod_a, cod_b, cod_c, cod_d)
{
		document.location.href('protocolo.asp?tipo=relatorio&acao=redefinir');
}

</SCRIPT>




</body>
