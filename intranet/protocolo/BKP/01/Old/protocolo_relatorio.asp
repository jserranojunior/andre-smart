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
selecao_top = request("selecao_top")
ok = request("ok")
data = request("data")
intervalo = request("intervalo")
campos = request("campos")
	if data = 1 Then
		campos = intervalo&campos
	end if
materiais = request("materiais")
subtotais = request("subtotais")
emprestimos = request("emprestimos")

cd_unidade = request("cd_unidade")
nm_unidade = request("nm_unidade")
	nm_und = request("nm_und")
	if nm_und <> "" Then
	nm_unidade = nm_und
	campos = campos&",a053_codung"
	end if

nm_protocolo = request("nm_protocolo")
	if nm_protocolo = "" Then
	nm_protocolo = ""
	elseif IsNumeric(nm_protocolo) then
 	nm_protocolo = int(nm_protocolo)
	end if

nm_paciente = request("nm_paciente")
a002_pacida_i = request("a002_pacida_i")
a002_pacida_f = request("a002_pacida_f")
cd_sexo = request("cd_sexo")
a054_codcon = request("a054_codcon")
	nm_convenio = request("nm_convenio")
	
nm_medico = request("nm_medico")
cd_tipoagenda = request("cd_tipoagenda")
cd_rack = request("cd_rack")
cd_especialidade = request("cd_especialidade")

dia_atual = day(now)
mes_atual = month(now)
ano_atual = year(now)
hora_atual = hour(now)
minuto_i = minute(now)
		'***			
dia_i = request("dia_i")
mes_i = request("mes_i")
ano_i = request("ano_i")
hora_i = request("hora_i")
minuto_i = request("minuto_i")
		'***			
dia_f = request("dia_f")
mes_f = request("mes_f")
ano_f = request("ano_f")
hora_f = request("hora_f")
minuto_f = request("minuto_f")



'**********************************************
'** Verifica quais campos foram selecionados **
'**********************************************
		'campo_0_1 = Instr(1,campos,"a053_codung",0)
		campo_1_1 = Instr(1,campos,"a053_codung",0)
		campo_2_1 = Instr(1,campos,"a002_numpro",0)
		campo_3_1 = Instr(1,campos,"a002_pacnom",0)
		campo_4_1 = Instr(1,campos,"a002_pacida",0)
		campo_5_1 = Instr(1,campos,"a002_pacsex",0)
		campo_6_1 = Instr(1,campos,"a054_codcon",0)
		campo_7_1 = Instr(1,campos,"a055_desmed",0)
		'campo_8_1 = Instr(1,campos," ",0)
		campo_9_1 = Instr(1,intervalo,"dt_ano, dt_mes",0)
		campo_9_2 = Instr(1,intervalo,"a002_datpro",0)
		campo_10_1 = Instr(1,campos,"a002_horage",0)
		campo_11_1 = Instr(1,campos,"a056_codrac",0)
		campo_12_1 = Instr(1,campos,"a057_codesp",0)
		campo_13_1 = Instr(1,campos,"a002_tipage",0)
		campo_14_1 = Instr(1,campos,"a058_codpro",0)
		campo_15_1 = Instr(1,materiais,"materiais",0)

'******************** Tratamento de dados************************
'-------------------- Período --------------------
	if ano_i = "" Then 'Caso o usuário apague o ano
		ano_i = ano_atual
	end if
	if ano_f = "" Then 'Caso o usuário apague o ano
		ano_f = ano_atual
	end if
	
	if ultimodiames(mes_f,ano_f) < int(dia_f) then
		dia_f = ultimodiames(mes_f,ano_f)
	end if
		'***
		data_i = mes_i&"/"&dia_i&"/"&ano_i
		data_f = mes_f&"/"&dia_f&"/"&ano_f
		'***
		
		if campo_9_1 > 0 then 'Mensal
		'if data > 0 then 'Mensal
			ultimo_dia_f = ultimodiames(mes_f,ano_f) 	'ultimo_dia_i = ultimodiames(mes_i,ano_i)
			
			If ano_i&zero(mes_i) <= ano_f&zero(mes_f) Then '1 - Menor ou igual
			periodo = "A002_datpro BETWEEN '"&mes_i&"/1/"&ano_i&"' AND '"&mes_f&"/"&ultimo_dia_f&"/"&ano_f&"'"
			elseIf ano_i&zero(mes_i)&zero(dia_i) > ano_f&zero(mes_f)&zero(dia_f) Then '1 - Maior
			'periodo = "A002_datpro BETWEEN '"&mes_i&"/1/"&ano_i&"' AND '"&mes_i&"/"&ultimo_dia_i&"/"&ano_i&"'"
			ok = ""
			periodo_erro = "<b>O período selecionado é inválido!</b> A data inicial é maior que a final.<br>&nbsp;"
			end if
			'response.write("mensal")
		Elseif campo_9_2 > 0 then 'diario
		'Elseif data > 0 then 'diario
			If ano_i&zero(mes_i)&zero(dia_i) <= ano_f&zero(mes_f)&zero(dia_f) Then '1 - Menor ou igual
			periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_f&"'"
			elseIf ano_i&zero(mes_i)&zero(dia_i) > ano_f&zero(mes_f)&zero(dia_f) Then '1 - Maior
			'periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_i&"'"
			ok = ""
			periodo_erro = "<b>O período selecionado é inválido!</b> A data inicial é maior que a final.<br>&nbsp;"
			end if
			'response.write("diário")
		End if

'-------------------- Unidade --------------------
	if nm_unidade <> ""  Then
		nm_unidade = int(nm_unidade)
		
		'Verifica as unidades solicitadas e monta o comando para o SQL
		numero = 1
		proc_array = split(nm_unidade,",")
		For Each proc_item In proc_array
				if numero = 1 Then
				nome_unidade = proc_item
				elseif numero >= 2 then
				'nome_unidade = nome_unidade&"%' or nm_unidade like '%"&proc_item
				nome_unidade = nome_unidade&"%' or A053_codung like '%"&proc_item
				end if
		numero = numero + 1
		Next
		'---------------------------------------------------------
			if not IsNumeric(nm_unidade) then
				'nome_unidade = " AND nm_unidade like '%"&nome_unidade&"%'"
				nome_unidade = " AND A053_desung like '%"&nome_unidade&"%'"
				campos = campos&",A053_desung, nm_sigla"
					cd_unidade = 0
					'response.write("não_numérico")
			else
				nome_unidade = " AND a053_codung like '"&nome_unidade&"'"
				'campos = campos&",nm_unidade, nm_sigla"
				campos = campos&",A053_desung, nm_sigla"
					cd_unidade = 0
					'response.write("numérico")
			end if
			
	Elseif cd_unidade > 0 then
		unidade = "AND a053_codung = "&cd_unidade&""
		campos = campos&",nm_sigla"
	end if

'-------------------- Protocolo --------------------
	if nm_protocolo <> "" then
		protocolo = "AND a002_numpro like '%"&nm_protocolo&"%'"
		campos = campos&",a002_numpro"
	end if

'-------------------- Paciente --------------------
	if nm_paciente <> "" then
		'if not nm_paciente IsNumeric then
		if not IsNumeric(nm_paciente) then
			paciente = "AND a002_pacnom like '"&nm_paciente&"%'"
			campos = campos&",a002_pacreg,a002_pacnom"
		Else
			paciente = "AND a002_pacreg like '"&nm_paciente&"%'"
			campos = campos&",a002_pacreg,a002_pacnom"
		end if
	end if

'-------------------- Idade --------------------
	if a002_pacida_i <> "" AND a002_pacida_f <> "" Then
		idade = "AND a002_pacida BETWEEN "&a002_pacida_i&" AND "&a002_pacida_f&""
	Elseif a002_pacida_i <> "" AND a002_pacida_f = "" Then
		idade = "AND a002_pacida BETWEEN "&a002_pacida_i&" AND "&a002_pacida_i&""
	end if

'-------------------- Sexo --------------------
	if cd_sexo <> "" then
		sexo = "AND a002_pacsex='"&cd_sexo&"'"
	end if

'-------------------- Convenio --------------------
	if nm_convenio <> "" Then
		if not IsNumeric(nm_convenio) then
			convenio = "AND a054_descon LIKE '%"&nm_convenio&"%'"
			campos = campos&",a054_codcon"
		Else
			convenio = "AND a054_codcon = "&nm_convenio&""
			campos = campos&", a054_codcon"
		End if
	end if

'-------------------- Médico --------------------
	'Verifica os cirurgiões solicitados e monta o comando para o SQL
		
	if nm_medico <> "" then
		if not IsNumeric(nm_medico) then
			medico = "AND a055_desmed like '%"&nm_medico&"%'"
			campos = campos&", a055_desmed, a055_crmmed"
		Else
			medico = "AND cd_crm = "&nm_medico&""
			campos = campos&", a055_crmmed, a055_desmed "
		end if
	end if

'-------------------- Tipo Agenda --------------------
	if campo_13_1 > 0 And cd_tipoagenda <> "" Then
		tipoagenda = "AND a002_tipage = '"&cd_tipoagenda&"'"
	end if
		if emprestimos = 1 Then
			'exclui_E = "AND (a002_tipage <> 'e')"
		tipoagenda = "AND a002_tipage <> 'E' "' "
		end if

'-------------------- Hora Agenda --------------------
		'***
		agenda_i = hora_i&":"&minuto_i
		agenda_f = hora_f&":"&minuto_f
		'***
		
		if campo_10_1 > 0 AND hora_i <> "" then
			If zero(hora_i)&zero(minuto_i) <= zero(hora_f)&zero(minuto_f) Then '1 - Menor ou igual
				agenda = "AND ({ fn hour(A002_horage) }) BETWEEN '"&hora_i&"' AND  '"&hora_f&"' AND ({ fn minute(A002_horage) }) BETWEEN '"&minuto_i&"' AND '"&minuto_f&"'" 
			elseIf zero(hora_i)&zero(minuto_i) > zero(hora_f)&zero(minuto_f) Then '1 - Maior
				ok = ""
				agenda_erro = "<b>O período selecionado é inválido!</b> O horário inicial é maior que o final.<br>&nbsp;"
			end if
		End if

'-------------------- Rack --------------------
		if campo_11_1 > 0 Then
			if cd_rack <> "" Then
				rack = "AND a056_codrac = '"&cd_rack&"'"
			else
				rack = ""
			end if
			campos = campos&", a056_desrac"
		end if
		
'-------------------- Especialidade --------------------
		if campo_12_1 > 0 Then
			if cd_especialidade > "0" Then
				especialidade = "AND a057_codesp = '"&cd_especialidade&"'"
			end if
			campos = campos&", a057_desesp"
		end if

'-------------------- Procedimentos --------------------
		if campo_14_1 > 0 Then
			'campos = campos&", a002_numseq"
			'campos = campos&", a058_codpro"
		end if

'-------------------- Materiais ------------------------
		if campo_15_1 > 0 Then
			campos = campos&", a002_numseq"
		end if

'**********************************************
'************** Seleção ***********************
'**********************************************
'selecao = " COUNT (a002_numpro) AS conta"
if selecao_top <> "" Then
	selecao_TOP = "TOP "&selecao_top
end if

objetivo = "a002_numpro"
'objetivo = "a058_codpro"

selecao = selecao_top&" COUNT ("&objetivo&") AS conta"

if campos = "" Then
selecao = selecao
Else
selecao = selecao &", "&campos 
End if
	'Elimina uma possível duplicação de virgulas
	selecao = replace(selecao,", ,",",")

'**********************************************
'************** Condições *********************
condicao = "Where "&periodo&" "&unidade&" "&nome_unidade&" "&protocolo&" "&paciente&" "&idade&" "&sexo&" "&convenio&" "&medico&" "&agenda&" "&tipoagenda&" "&rack&" "&especialidade&""

'**********************************************
'************** Agrupamento *********************
if campos <> "" Then
agrupamento = "Group by "&campos
	teste_agrupamento = Instr(1,agrupamento,"Group by ,",0)
			if teste_agrupamento = 1 then
			agrupamento = replace(agrupamento,"Group by ,","Group by ")'Retira a virgula
			end if
			agrupamento = replace(agrupamento,", ,",",  ")'Retira a virgula dupla
			
			'response.write(teste_agrupamento)

'**********************************************
'**************** Ordem ***********************
sentido = request("sentido")

ordem_qtd = request("ordem_qtd")'0
ordem_unidade = request("ordem_unidade")'1
ordem_protocolo = request("ordem_protocolo")'2
ordem_paciente = request("ordem_paciente")'3
ordem_idade = request("ordem_idade")'4
ordem_sexo = request("ordem_sexo")'5
ordem_convenio = request("ordem_convenio")'6
ordem_medico = request("ordem_medico")'7

ordem_data = request("ordem_data")'9
ordem_hragenda = request("ordem_hragenda")'10
ordem_rack = request("ordem_rack")'11
ordem_especialidade = request("ordem_especialidade")'12
ordem_tipoagenda = request("ordem_tipoagenda")'13
ordem_procedimento = request("ordem_procedimento")'14
ordem_materiais = request("ordem_materiais")'15
%>

<!--#include file="../includes/ordem_protocolos.asp"-->
<%
	if ordem <> "" Then
		ordem = "ORDER BY "&ordem
			ordem = replace(ordem,"ORDER BY ,","ORDER BY ")'Retira a virgula indevida
				Else
	ordem = ""
	end if

end if%>

<body>
<!--table align="center" border="1">
	<tr>
		<td align="center" bgcolor="#C0C0C0"><b>RELATÓRIOS - <%'=cd_patrimonio_vazio%><font color="red">Protocolos</font></b>
		</td>
	</tr>
		<tr><td align="center"-->
			<table border="1" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" id="no_print">
				<form action="protocolo.asp" method="post" name="form">
				<input type="hidden" name="tipo" value="relatorio">
				<tr id="no_print">
					<td bgcolor="#C0C0C0" align="center" colspan="100%"><b>RELATÓRIOS - <font color="red">Protocolos</font></b></td>
				</tr>
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
<!-- *** Data *** -->
				<tr id="no_print" bgcolor="#e6e6e6">
					<td colspan="3">
						<%if data <> 0  then%>
					<%ck_data = "checked"%>
					<%end if%>						
						<input type="checkbox" name="data" value="1" <%=ck_data%>>&nbsp;Data	</td>
					<td colspan="4"><%if campo_9_2 > 0 then%><%ck_campo_9_2 ="checked"%><%else%><%ck_campo_9_1 ="checked"%><%end if%>
						<input type="radio" name="intervalo" value="dt_ano, dt_mes" <%=ck_campo_9_1%>>mensal
						<input type="radio" name="intervalo" value="a002_datpro" <%=ck_campo_9_2%>>diário
						&nbsp;&nbsp;&nbsp;
						<select name="dia_i" class="inputs">
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
						<select name="mes_i" class="inputs">
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
						<input type="text" name="ano_i" size="4" maxlength="4" value="<%=ano_i%>" class="inputs">
						até
						<select name="dia_f" class="inputs">
							<%numero = 1
							do while NOT numero = 32
								if int(dia_f) = numero Then
								check = "selected"
								end if%>					
							<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
						</select>/
						<select name="mes_f" class="inputs">
							<%numero = 1
							do while NOT numero = 13
								if mes_f = "" Then
									mes_f = 12
									if mes_hoje = numero Then
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
						<input type="text" name="ano_f" size="4" maxlength="4" value="<%=ano_f%>" class="inputs"> 
					</td>
					<td><select name="ordem_data" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
<!-- *** Quantidade *** -->
				<tr id="no_print">
					<td colspan="3">&nbsp;<!--<b>TOP: <select name="selecao_top">
												<option value="">Todos</option>
												<option value="10">10 maiores</option>
												<option value="50">50 maiores</option>
											</select>--></b></td>
					<td colspan="4">&nbsp;<!--<b>Quantidade:</b> &nbsp;<input type="radio" name="objetivo" value="a058_codpro">Protocolos&nbsp; - &nbsp;
									<input type="radio" name="objetivo" value="a058_codpro">Especialidades&nbsp;--><!-- - &nbsp;
									<input type="radio" name="objetivo" value="a058_codpro">Materiais-->
					</td>
					<td><select name="ordem_qtd" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
					<td colspan="3">
					<%if campo_1_1 <> 0 OR cd_unidade <> 0 OR nm_unidade <> "" then%>
					<%ck_campo_1 = "checked"%>
					<%end if%>
						<input type="checkbox" name="campos" value=",a053_codung, nm_sigla" <%=ck_campo_1%>>&nbsp;Unidade 
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
						do while NOT numero = 16
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
					<td colspan="3"><%if campo_2_1 <> 0 then%><%ck_campo_2 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_numpro" <%=ck_campo_2%>>&nbsp;Protocolo						
					</td>
					<td colspan="4"><input type="text" name="nm_und" size="2" maxlength="3" class="inputs" value="<%=nm_und%>" class="inputs">.<input type="text" name="nm_protocolo" size="10" maxlength="50" class="inputs" value="<%=mm_protocolo%>" class="inputs">
						
					</td>
					<td><select name="ordem_protocolo" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
<!-- *** Paciente *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_3_1 <> 0 OR nm_paciente <> "" then%><%ck_campo_3 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_pacreg, a002_pacnom" <%=ck_campo_3%>>&nbsp;Paciente						
					</td>
					<td colspan="4"><input type="text" name="nm_paciente" size="50" maxlength="500" value="<%=nm_paciente%>" class="inputs"></td>
					<td><select name="ordem_paciente" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
					<td colspan="3"><%if campo_4_1 <> 0 then%><%ck_campo_4 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_pacida" <%=ck_campo_4%>>&nbsp;Idade
					</td>
					<td colspan="4">				
						<input type="text" size="3" name="a002_pacida_i" maxlength="3" value="<%=a002_pacida_i%>" class="inputs"> até 
						<input type="text" size="3" name="a002_pacida_f" maxlength="3" value="<%=a002_pacida_f%>" class="inputs">
					</td>
					<td><select name="ordem_idade" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
					<td colspan="3"><%if campo_5_1 <> 0 then%><%ck_campo_5 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_pacsex" <%=ck_campo_5%>>&nbsp;Sexo
					</td>
					<td colspan="4"><%if cd_sexo = "F" Then%><%ck_fem = "selected"%><%Elseif cd_sexo = "M" Then%><%ck_masc = "selected"%><%end if%>
					<select name="cd_sexo" size="1" class="inputs">
						<option value=""></option>
						<option value="F" <%=ck_fem%>>F</option>
						<option value="M" <%=ck_masc%>>M</option>
					</select>
					</td>
					<td><select name="ordem_sexo" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
					<td colspan="3"><%if campo_6_1 <> 0 then%><%ck_campo_6 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a054_codcon, a054_descon" <%=ck_campo_6%>>&nbsp;Convênio 
					</td>
					<td colspan="4"><input type="text" name="nm_convenio" size="50" maxlength="500" value="<%=nm_convenio%>" class="inputs"></td>
					<td><select name="ordem_convenio" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
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
<!-- *** Médico *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_7_1 <> 0 OR a055_desmed <> "" then%><%ck_campo_7 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a055_crmmed, a055_desmed" <%=ck_campo_7%>>&nbsp;Médico
					</td>
					<td colspan="4"><input type="text" name="nm_medico" size="50" maxlength="500" value="<%=nm_medico%>" class="inputs">
						
					</td>
					<td><select name="ordem_medico" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_medico) = numero Then
							check = "selected"
						end if
						%>
						<option value="<%=numero%>" <%=check%>><%=numero%></option>
						<%numero = numero+1
						check = ""
						loop%>
					</select>
					</td>
<!-- *** Hora Agenda *** -->
				<tr id="no_print">
					<td colspan="3">
						<%if campo_10_1 > 0 then%><%ck_campo_10 ="checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_horage" <%=ck_campo_10%>>&nbsp;Hora Agenda
					</td>
					<td colspan="4">&nbsp;
					<!--	<input type="text" name="hora_i" value="<%=hora_i%>" size="2" maxlength="2">:<input type="text" name="minuto_i" value="<%=minuto_i%>" size="2" maxlength="2"> Até
						<input type="text" name="hora_f" value="<%=hora_f%>" size="2" maxlength="2">:<input type="text" name="minuto_f" value="<%=minuto_f%>" size="2" maxlength="2">
					-->
					</td>
					<td><select name="ordem_hragenda" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_hragenda) = numero Then
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
<!-- *** Tipo de agenda *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_13_1 <> 0 then%><%ck_campo_13 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_tipage" <%=ck_campo_13%>>&nbsp;Tipo Agenda						
					</td>
					<td colspan="4">
						<select name="cd_tipoagenda" class="inputs"> 
							<option value=""></option>
							<option value="N" <%if cd_tipoagenda = "N" Then%><%="SELECTED"%><%end if%>>Normal</option>
							<option value="A" <%if cd_tipoagenda = "A" Then%><%="SELECTED"%><%end if%>>A seguir</option>
							<option value="E" <%if cd_tipoagenda = "E" Then%><%="SELECTED"%><%end if%>>Empréstimo</option>
							<option value="U" <%if cd_tipoagenda = "U" Then%><%="SELECTED"%><%end if%>>Urgente</option>
							<option value="P" <%if cd_tipoagenda = "P" Then%><%="SELECTED"%><%end if%>>Pós Agendada</option>
						</select>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if emprestimos = 1 Then%><%ck_emprestimos = "checked"%><%end if%>
						<input type="checkbox" name="emprestimos" value="1" <%=ck_emprestimos%>>&nbsp;sem emprestimos
					</td>
					<td><select name="ordem_tipoagenda" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_tipoagenda) = numero Then
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
<!-- *** Rack *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_11_1 <> 0 then%><%ck_campo_11 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a056_codrac, a002_numseq" <%=ck_campo_11%>>&nbsp;Rack
					</td>
					<td colspan="4">
						<select name="cd_rack" class="inputs"> 
							<option value=""></option>
							<%if cd_rack = "" Then%><%cd_rack = 0%><%end if%>
							
							<%strsql_rack = "SELECT * FROM TBL_rack"
							Set rs_rack = dbconn.execute(strsql_rack)
							
							while not rs_rack.EOF
							cod_rack = rs_rack("a056_codrac")
							nm_rack = rs_rack("a056_desrac")
							
							if cod_rack = int(cd_rack) Then
							ck_rack = "selected"
							else
							ck_rack = ""
							end if%>						
							<option value="<%=cod_rack%>" <%=ck_rack%>> <%=nm_rack%></option>
							<%rs_rack.movenext
							ck_rack = ""
							wend%>
						</select>
					</td>
					<td><select name="ordem_rack" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_rack) = numero Then
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
<%if xpto = 1 then%>
<!-- *** Especialidade *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_12_1 <> 0 then%><%ck_campo_12 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a057_codesp" <%=ck_campo_12%>>&nbsp;Especialidade						
					</td>
					<td colspan="4">
						<select name="cd_especialidade" class="inputs"> 
							<option value="0"></option>
							<%strsql_esp = "SELECT * FROM TBL_especialidades Order by 'nm_especialidade'"
							Set rs_esp = dbconn.execute(strsql_esp)
							
							while not rs_esp.EOF
							cd_esp = rs_esp("cd_codigo")
							nm_especialidade = rs_esp("nm_especialidade")
							
							if int(cd_especialidade) = cd_esp Then
							ck_esp = "selected"
							end if%>						
							<option value="<%=cd_esp%>" <%=ck_esp%>><%=nm_especialidade%></option>
							<%rs_esp.movenext
							ck_esp = ""
							wend%>
						</select>
					</td>
					<td><select name="ordem_especialidade" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_especialidade) = numero Then
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
<!-- *** Procedimento *** -->
				<tr id="no_print">
					<td colspan="3"><%if campo_14_1 <> 0 then%><%ck_campo_14 = "checked"%><%end if%>
						<!--input type="checkbox" name="campos" value=",a058_codpro,a002_numseq" <%=ck_campo_14%>-->
						<input type="checkbox" name="campos" value=",a058_codpro" <%=ck_campo_14%>>&nbsp;Procedimento
					</td>
					<td colspan="4"><input type="text" name="procedimento" size="50" maxlength="50" class="inputs">
						
					</td>
					<td><select name="ordem_procedimento" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_procedimento) = numero Then
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
<!-- *** Materiais *** -->
				<!--tr id="no_print">
					<td colspan="3"><%if campo_15_1 <> 0 then%><%ck_campo_15 = "checked"%><%end if%>
						<input type="checkbox" name="materiais" value="materiais" <%=ck_campo_15%>>&nbsp;Materiais
					</td>
					<td colspan="4">&nbsp;
						<!--select name="cd_materiais">
							<option value="0"></option>
							
						</select-->
					</td>
					<td><select name="ordem_materiais" size="1">
						<option value="0"></option>
						<%numero = 1
						do while NOT numero = 16
						if int(ordem_materiais) = numero Then
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
<%end if%>
<!-- *** Subtotal-2 ***-->
				<tr><%if subtotais > 0 then%><%ck_subtotais = "checked"%><%end if%>
					<td colspan="3"><input type="checkbox" name="subtotais" value="1" <%=ck_subtotais%>>&nbsp;Mostrar subtotais.</td>
				</tr>
				<tr>
					<td colspan="3">.</td>
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
						<!--input type="hidden" name="x" value="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->
						<input type="submit" name="ok" value="Gerar Relatório">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input onclick="javascript:JsRedefinir('')" type="button" value="Limpar" title="Redefinir"> 
					</td>
				</tr>
				<tr class="txt">
					<td align="left" colspan="100%">
					<font color="red"><%=periodo_erro%></font><br>
					</td>
				</tr>
				<!--tr id="no_print">
					<td align="left" colspan="100%" bgcolor="#ffff00">
					<b>SELECT </b><%=selecao%><br>
					<b>FROM VIEW_protocolo_lista </b><br>
					<%=condicao%> <%=cond_periodo%> <%=cond_und%> <br>
					<%=agrupamento%><br>
					<%=ordem%> &nbsp;<%=sentido%><br>
					</td>
				</tr-->
			</table>
			<br><br>
			
			
			<table align="left" border="0" cellpadding="1" cellspacing="1">
			</form>
	<%if ok <> "" Then 'Inicio do resultado da pesquisa
			'******************************
			'*							  *
			'* 	  2ª Parte - Resultados   *
			'*                            *
			'******************************	%>
				<tr id="ok_print">
					<td bgcolor="#C0C0C0" align="center" colspan="100%"><b>RELATÓRIOS - <font color="red">Protocolos</font></b></td>
				</tr>
				
				<tr bgcolor="#c0c0c0">
					<td>Qtd<br><img src="../imagens/blackdot.gif" alt="" width="22" height="1" border="0"></td>
					<%if campo_1_1 > 0 OR cd_unidade > 0 OR nm_unidade <> "" then%><td>Unidade<br><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td><%end if%>
					<%if campo_2_1 > 0 then%><td>Protocolo<br><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td><%end if%>
					<%if campo_3_1 > 0 OR nm_paciente <> "" then%><td colspan="2" align>Paciente<br><img src="../imagens/blackdot.gif" alt="" width="200" height="1" border="0"></td><%end if%>
					<%if campo_4_1 > 0 then%><td>Idade<br><img src="../imagens/blackdot.gif" alt="" width="35" height="1" border="0"></td><%end if%>
					<%if campo_5_1 > 0 then%><td>Sexo<br><img src="../imagens/blackdot.gif" alt="" width="30" height="1" border="0"></td><%end if%>
					<%if campo_6_1 > 0 then%><td>Convênio<br><img src="../imagens/blackdot.gif" alt="" width="110" height="1" border="0"></td><%end if%>
					<%if campo_7_1 > 0 then%><td colspan="2">Equipe Médica<br><img src="../imagens/blackdot.gif" alt="" width="200" height="1" border="0"></td><%end if%>
					<%if data > 0 then%><td>Data<br><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td><%end if%>
					<%if campo_10_1 > 0 then%><td>Agenda<br><img src="../imagens/blackdot.gif" alt="" width="45" height="1" border="0"></td><%end if%>
					<%if campo_11_1 > 0 then%><td>Rack<br><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td><%end if%>
					<%if campo_12_1 > 0 then%><td>Especialidade<br><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td><%end if%>
					<%if campo_13_1 > 0 then%><td>tipo<br><img src="../imagens/blackdot.gif" alt="" width="50" height="1" border="0"></td><%end if%>
					<%if campo_14_1 > 0 then%><td>Procedimento<br><img src="../imagens/blackdot.gif" alt="" width="250" height="1" border="0"></td><%end if%>
					<%if campo_15_1 > 0 then%><td>Materiais<br><img src="../imagens/blackdot.gif" alt="" width="150" height="1" border="0"></td><%end if%>
					<%'if campo_2_1 > 0 then%><!--td id="no_print"></td--><%'end if%>
				</tr>
<%'xsql = "SELECT count(a002_numpro)as pre_conta FROM VIEW_protocolo_lista "&condicao&" "&agrupamento&""

xsql = "SELECT "&selecao&" FROM VIEW_protocolo_lista "&condicao&" "&agrupamento&" "&ordem&" "&sentido&""
Set rs_pc = dbconn.execute(xsql)
linhas = 0
do while Not rs_pc.EOF
	linhas = 1
rs_pc.movenext
pre_conta = linhas + pre_conta
loop%>
<%if pre_conta > 500 Then%>
	<tr><td colspan="100%">A busca retornou muitas linhas, por favor refine a sua busca, para . (<%=linhas%>/<%=pre_conta%>)</td></tr>			
<%else%>
				<tr><td colspan="100%" id="ok_print"><img src="imagens/blackdot.gif" alt="" width="100%" height="1"></td></tr>
	<%	
			principal = primeiro
			nr = 0
			
			xsql = "SELECT "&selecao&" FROM VIEW_protocolo_lista "&condicao&" "&agrupamento&" "&ordem&" "&sentido&" "
			Set rs = dbconn.execute(xsql)
			
			do while Not rs.EOF
			
			conta = rs("conta")
			
			
			if campo_1_1 > 0 OR cd_unidade > 0 OR nm_unidade <> "" then
				cd_unidade = rs("a053_codung")
					cd_unidade = rs("a053_codung")
					a053_codung = cd_unidade&"."
				nm_und = rs("nm_sigla")
					if ordem_unidade = 1 Then
					compara = a053_codung
					end if
			end if
			
			if campo_2_1 > 0 then
				'a053_codung = rs("a053_codung")
				a002_numpro = rs("a002_numpro")
					if ordem_protocolo = 1 Then
					compara = a002_numpro
					end if
			end if
			
			if campo_3_1 > 0 OR nm_paciente <> "" then
				a002_pacreg = rs("a002_pacreg")
				a002_pacnom = rs("a002_pacnom")
					if ordem_paciente = 1 Then
					compara = a002_pacreg
					end if
			end if
			
			if campo_4_1 > 0 then
				a002_pacida = rs("a002_pacida")
					if ordem_idade = 1 Then
					compara = a002_pacida
					end if
			end if
			
			if campo_5_1 > 0 then
				a002_pacsex = rs("a002_pacsex")
					if a002_pacsex = "M" Then
					sexo = "Masc"
					Elseif a002_pacsex = "F" Then
					sexo = "Fem"
					End if
						if ordem_sexo = 1 Then
						compara = a002_pacsex
						end if
			end if
			
			if campo_6_1 > 0 then
				'a054_codcon = rs("a054_codcon")
				a054_descon = rs("a054_descon")
					if ordem_convenio = 1 Then
					compara = a054_descon
					end if
			end if
			
			if campo_7_1 > 0 then
				cirurgiao = rs("a055_desmed")
				crm = rs("a055_crmmed")
					if ordem_medico = 1 Then
					compara = cirurgiao
					end if
			end if
			
			'Não há campo_8 ainda
			
			if data = 1 AND campo_9_1 > 0 then
				dt_mes = rs("dt_mes")
				dt_ano = rs("dt_ano")
				A002_datpro = left(mes_selecionado(dt_mes),3)&"/"&dt_ano
					if ordem_data = 1 Then
					compara = A002_datpro
					end if
			Elseif data = 1 AND campo_9_2 > 0 then
				A002_datpro_dia = day(rs("A002_datpro"))
				A002_datpro_mes = month(rs("A002_datpro"))
				A002_datpro_ano = year(rs("A002_datpro"))
					A002_datpro = zero(A002_datpro_dia)&"/"&zero(A002_datpro_mes)&"/"&A002_datpro_ano
					if ordem_data = 1 Then
					compara = A002_datpro
					end if
			end if
			
			if campo_10_1 > 0 then
				agenda_hr = zero(hour(rs("a002_horage")))&":"&zero(minute(rs("a002_horage")))
					if ordem_hragenda = 1 Then
					compara = agenda_hr
					end if
			end if
			
			if campo_11_1 > 0 then
				a056_codrac = rs("a056_codrac")
				a056_desrac = rs("a056_desrac")
					if ordem_rack = 1 Then
					compara = a056_desrac
					end if
			end if
			
			if campo_12_1 > 0 then
				'a057_codesp = rs("a057_codesp")
				'a057_desesp = rs("a057_desesp")
				a057_codesp = rs("cd_codigo")
				a057_desesp = rs("nm_especialidade")
					if ordem_especialidade = 1 Then
					compara = nm_especialidade
					end if
			end if
			
			if campo_13_1 > 0 then
				a002_tipage = rs("a002_tipage")
				if a002_tipage = "A" then
					tipo_agenda = "A seguir"
				Elseif a002_tipage = "E" then
					tipo_agenda = "Empréstimo"
				Elseif a002_tipage = "N" then
					tipo_agenda = "Normal"
				Elseif a002_tipage = "P" then
					tipo_agenda = "Pós agendada"
				Elseif a002_tipage = "U" then
					tipo_agenda = "Urgente"
				else
					tipo_agenda = ""
				end if
					if ordem_tipoagenda = 1 Then
					compara = a002_tipage
					end if
			end if
			
			if campo_14_1 > 0 then
				'cd_procedimento = rs("a002_numseq")
				a058_codpro = rs("a058_codpro")
				'nome_especialidade = rs("nm_especialidade")
			end if
			
			if campo_15_1 > 0 then
				cd_protocolo = rs("a002_numseq")
				'a058_codpro = rs("a058_codpro")
				'nome_especialidade = rs("nm_especialidade")
			end if%>
			
				<%if grava_info <> compara AND nr > 0 AND subtotais > 0 Then%>
				<tr><td bgcolor="#ebebeb" colspan="100%"><b><i><%=sub_conta%></i></b></td></tr><tr><td>&nbsp;</td></tr>
				<%sub_conta = 0
				end if%>
				<tr onmouseover="mOvr(this,'#CFC8FF');"  onclick="javascript:;" onmouseout="mOut(this,'');">				
					<td valign="top"><%=zero(conta)%></td>
				<%if campo_1_1 > 0 OR cd_unidade > 0 OR nm_unidade <> "" then%>
					<td valign="top"><%=nm_und%></td><%end if%>
				<%if campo_2_1 > 0 then%>	
					<td valign="top"><%if cd_unidade <> "" AND a002_numpro <> "" Then%>
					<a href="javascript: void(0);" onclick="window.open('protocolo/protocolo_folha.asp?tipo=ver&cd_unidade=<%=cd_unidade%>&cd_protocolo=<%=proto(a002_numpro)%>&cd_digito=<%=right(digitov(cdun(a053_codung)&proto(a002_numpro)),2)%>&janelas_cadastros', 'principal', 'width=430, height=570'); return false;">
					<%end if%>
					<%=digitov(cdun(a053_codung)&proto(a002_numpro))%></a><!--input type="checkbox" name="nada" value=""--></td><%end if%>			
				<%if campo_3_1 > 0 OR nm_paciente <> "" then%>
					<td valign="top"><%=a002_pacreg%></td><td valign="top"><%=a002_pacnom%></td><%end if%>
				<%if campo_4_1 > 0 then%>
					<td valign="top"><%=a002_pacida%></td><%end if%>
				<%if campo_5_1 > 0 then%>
					<td valign="top"><%=sexo%></td><%end if%>
				<%if campo_6_1 > 0 then%>
					<td valign="top"><%=a054_codcon%> &nbsp; <%=a054_descon%></td><%end if%>
				<%if campo_7_1 > 0 then%>
					<td valign="top"><%=crm%></td><td valign="top">Dr. <%=cirurgiao%></td><%end if%>
				<%if data > 0 then%>
					<td valign="top"><%=A002_datpro%></td><%end if%>
				<%if campo_10_1 > 0 then%>
					<td valign="top"><%=agenda_hr%></td><%end if%>
				<%if campo_11_1 > 0 then%>
					<td valign="top"><%=a056_desrac%></td><%end if%>
				<%if campo_12_1 > 0 then%>
					<td valign="top"><%=a057_desesp%></td><%end if%>	
				<%if campo_13_1 > 0 then%>
					<td valign="top"><%=tipo_agenda%></td><%end if%>
				<%if campo_14_1 > 0 then%>	
					<td valign="top" width="800">
						<%
						if objetivo = "a058_codpro" then
						strsql_proc = "SELECT * FROM TBL_procedimento WHERE a058_codpro = "&a058_codpro&""' ORDER BY 'a003_propri' desc"
						Set rs_proc = dbconn.execute(strsql_proc)%>
							<%while not rs_proc.EOF
							'cd_proc = rs_proc("a002_numseq")
							a058_codpro = rs_proc("a058_codpro")							
							nm_procedimento = rs_proc("a058_despro")
							'a003_propri = rs_proc("a003_propri")
							if a003_propri <> "S" then%><i><font color="#4D4D4D"><%end if%><b><%=a058_codpro%></b></font><%=lcase(a003_propri)%><%if a003_propri = "s" then%></i><%end if%> - <%=nm_procedimento%><br>							
							<%rs_proc.movenext
							wend
						Else
						strsql_proc = "SELECT * FROM view_protocolo_procedimento WHERE a002_numseq = "&cd_procedimento&" ORDER BY 'a003_propri' desc"
						Set rs_proc = dbconn.execute(strsql_proc)%>
							<%while not rs_proc.EOF
							cd_proc = rs_proc("a002_numseq")
							a058_codpro = rs_proc("a058_codpro")							
							nm_procedimento = rs_proc("a058_despro")
							a003_propri = rs_proc("a003_propri")
							if a003_propri <> "S" then%><i><font color="#4D4D4D"><%end if%><b><%=a058_codpro%></b></font><%=lcase(a003_propri)%><%if a003_propri = "s" then%></i><%end if%> - <%=nm_procedimento%><br>							
							<%rs_proc.movenext
							wend
						end if%>
							
					 </td>
				<%end if%>
				<%if campo_15_1 > 0 then
					a002_numseq = rs("a002_numseq")%>
					
					<td valign="top" width="800">
						<%strsql_mat = "SELECT * FROM view_protocolo_material WHERE a002_numseq = '"&cd_protocolo&"'"
						Set rs_mat = dbconn.execute(strsql_mat)%>
							<%while not rs_mat.EOF
							a061_nommat = rs_mat("a061_nommat")%>
							<%=a061_nommat%><br>							
							<%rs_mat.movenext
							wend%>					
					 </td>
				<%end if%>
				<%'if campo_2_1 > 0 then%>	
					<!--td valign="top" id="no_print"><%if cd_unidade <> "" AND a002_numpro <> "" Then%><%'response.write("<a href=protocolo.asp?tipo=dvd&cd_unidade="&cd_unidade&"&cd_protocolo="&a002_numpro&">")%><%end if%>
					<img src="../imagens/filme3.gif" alt="" width="10" height="11" border="0"></a></td-->
				<%'end if%>	
				
				</tr>
				
				
				
			<%rs.movenext
				grava_info = compara
				sub_conta = sub_conta + conta
			geral = geral + conta
			nr = 1
			loop
			
					
end if%>
			<%if subtotais = 1 then%>
				<tr><td bgcolor="#ebebeb" colspan="100%"><b><i><%=sub_conta%></i></b></td></tr><tr><td>&nbsp;</td></tr>
				<%end if%>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>					
				<tr bgcolor="#e1e1e1">
					<td colspan="100%">&nbsp;Total: <%=geral%><br><%=agenda_erro%></td>
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>
		<%End if%>
				
		
	<%'end if 'Fim do resultado da pesquisa%>	
		
		<tr bgcolor="#C0C0C0">
		<td align="left" colspan="100%" id="no_print">
		<!--
		<b>SELECT </b><%=selecao%><br>
		<b>FROM VIEW_protocolo_lista </b><br>
		<%=condicao%> <%=cond_periodo%> <%=cond_und%> <br>
		<%=agrupamento%><br>
		<%=ordem%><br>
		<br>
		<br>
		 Campos: <%=campos%><br>
		 o:<%=ordem%><br>
		 s:<%=sentido%><br><br>
		 (1-<%=primeiro%>), (2-<%=segundo%>), (3-<%=terceiro%>), (4-<%=quarto%>), (5-<%=quinto%>), (6-<%=sexto%>), (7-<%=setimo%>), (8<%=oitavo%>), (9-<%=nono%>), 10-(<%=decimo%>), (Fim-<%=fim%>), (U-<%=ultimo%>)
		<br>
		campo_9_1: (<%=campo_9_1%>) * <br>
		campo_9_2(<%=campo_9_2%>)<br>
		dia_f: (<%=dia_f%>)<br>
		Ok: <%=ok%><br>
		<%=nm_protocolo%>-->
		</td>
	</tr><%'end if%>
	<%'end if 'Fim do resultado da pesquisa%>

				<tr><td colspan="100%">&nbsp;</td></tr>
			</table>
		<!--/td><-- Fim da tabela principal- ->
	</tr>
</table--><br>
<br>

<SCRIPT LANGUAGE="javascript">

function JsRedefinir(cod_a, cod_b, cod_c, cod_d)
{
		document.location.href('protocolo.asp?tipo=relatorio&acao=redefinir');
}

</SCRIPT>




</body>
