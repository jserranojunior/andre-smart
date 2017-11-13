<html>
	
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
.txt_cinza {color: #6c6c6c;font-family: verdana;font-size:10px;text-decoration:none;}
.inputs {background-color: #cdcdcd; font: 12px verdana, arial, helvetica, sans-serif; color:#000000; border:0px solid #000000;}  
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
function ordem(a,ordem_total,obj){
	a=a;
	objeto=obj;
	ordem_total=ordem_total;
		if (ordem_total != ''){
			virg =  ',';
			}
		else{
		virg= ''
		}		
	ordem_total = ordem_total + virg
		if (a != ""){
			ordem_total = ordem_total + a;}
		
	document.form.ordem_res.value=a;
	document.form.ordem_total.value=(ordem_total.replace(",,", ","));		
	
	var el=document.getElementById(objeto);
	el.style.display=(el.style.display!='none'?'none':'');
	
	document.getElementsByName(objeto)[1].value= document.form.ordem_inicial.value+'º';
	document.form.ordem_inicial.value = document.form.ordem_inicial.value*1+1;
}
function limpa_ordem(obj){
	document.form.ordem_total.value='';
	
	document.form.ordem_1.value='';
	//document.form.ordem_2.value='';
	document.form.ordem_3.value='';
	document.form.ordem_4.value='';
	document.form.ordem_5.value='';
	//document.form.ordem_6.value='';
	//document.form.ordem_7.value='';
	//document.form.ordem_8.value='';
	document.form.ordem_9.value='';
	//document.form.ordem_10.value='';
	//document.form.ordem_11.value='';
	document.form.ordem_12.value='';
	document.form.ordem_13.value='';
	document.form.ordem_14.value='';
	//document.form.ordem_15.value='';
	/*document.form.ordem_16.value='';
	document.form.ordem_17.value='';*/
	
	document.form.ordem_inicial.value='1';
	
	document.ordem_1.style.display='inline';
	//document.ordem_2.style.display='inline';
	document.ordem_3.style.display='inline';
	document.ordem_4.style.display='inline';
	document.ordem_5.style.display='inline';
	//document.ordem_6.style.display='inline';
	//document.ordem_7.style.display='inline';
	//document.ordem_8.style.display='inline';
	document.ordem_9.style.display='inline';
	//document.ordem_10.style.display='inline';
	//document.ordem_11.style.display='inline';
	document.ordem_12.style.display='inline';
	document.ordem_13.style.display='inline';
	document.ordem_14.style.display='inline';
	//document.ordem_15.style.display='inline';
	/*document.ordem_16.style.display='inline';
	document.ordem_17.style.display='inline';*/
	
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
		campos = replace(campos,"so_ano","dt_ano")'*** Muda a palavra so_ano para dt_ano ***
	end if
	
tipo_relatorio = request("tipo_relatorio")
	if tipo_relatorio = 2 then
		campos = campos&""
			campos = replace(campos,"a053_codung","cd_unidade")
	elseif tipo_relatorio = 3 then
		campos = campos&""
	elseif tipo_relatorio = 4 then
		campos = campos&""
	end if

procedimentos = request("procedimentos")
str_procedimento = request("nm_procedimento")
materiais = request("materiais")
str_material = request("nm_material")
patrimonios = request("patrimonios")
str_patrimonio = request("nm_patrimonio")

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
cd_tipo_rack = request("cd_tipo_rack")
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
	if dia_f = "" Then
		dia_f = 31
	end if
mes_f = request("mes_f")
ano_f = request("ano_f")
hora_f = request("hora_f")
minuto_f = request("minuto_f")



'**********************************************
'** Verifica quais campos foram selecionados **
'**********************************************
		campo_1_1 = Instr(1,campos,"a053_codung",0)
		campo_2_1 = Instr(1,campos,"a002_numpro",0)
		campo_3_1 = Instr(1,campos,"a002_pacnom",0)
		campo_4_1 = Instr(1,campos,"a002_pacida",0)
		campo_5_1 = Instr(1,campos,"a002_pacsex",0)
		campo_6_1 = Instr(1,campos,"a054_codcon",0)
		campo_7_1 = Instr(1,campos,"a055_desmed",0)
		'campo_8_1 = Instr(1,campos," ",0)
		campo_9_0 = Instr(1,intervalo,"so_ano",0)
		campo_9_1 = Instr(1,intervalo,"dt_ano, dt_mes",0)
		campo_9_2 = Instr(1,intervalo,"a002_datpro",0)
		campo_10_1 = Instr(1,campos,"a002_horage",0)
		campo_10_2 = Instr(1,campos,"a002_horini,a002_horfin",0)
		campo_11_1 = Instr(1,campos,"a056_codrac",0)
		campo_11_2 = Instr(1,campos,"tipo_rack",0)
		campo_12_1 = Instr(1,campos,"a057_codesp",0)
		campo_13_1 = Instr(1,campos,"a002_tipage",0)
		campo_14_1 = Instr(1,procedimentos,"procedimentos,a002_numpro",0)
		campo_15_1 = Instr(1,materiais,"materiais",0)
		campo_16_1 = Instr(1,patrimonios,"patrimonios",0)

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
		
		If campo_9_0 > 0 then 'Mensal
			If ano_i <= ano_f Then '1 - Menor ou igual
			periodo = "A002_datpro BETWEEN '1/1/"&ano_i&"' AND '12/31/"&ano_f&"'"
			elseIf ano_i > ano_f Then '1 - Maior
			ok = ""
			periodo_erro = "<b>O período selecionado é inválido!</b> A data inicial é maior que a final.<br>&nbsp;"
			end if
		Elseif campo_9_1 > 0 then 'Mensal
			ultimo_dia_f = ultimodiames(mes_f,ano_f)
			
			If ano_i&zero(mes_i) <= ano_f&zero(mes_f) Then '1 - Menor ou igual
			periodo = "A002_datpro BETWEEN '"&mes_i&"/1/"&ano_i&"' AND '"&mes_f&"/"&ultimo_dia_f&"/"&ano_f&"'"
			elseIf ano_i&zero(mes_i)&zero(dia_i) > ano_f&zero(mes_f)&zero(dia_f) Then '1 - Maior
			ok = ""
			periodo_erro = "<b>O período selecionado é inválido!</b> A data inicial é maior que a final.<br>&nbsp;"
			end if
			'response.write("mensal")
		Elseif campo_9_2 > 0 then 'diario
			If ano_i&zero(mes_i)&zero(dia_i) <= ano_f&zero(mes_f)&zero(dia_f) Then '1 - Menor ou igual
			periodo = "A002_datpro BETWEEN '"&data_i&"' AND '"&data_f&"'"
			elseIf ano_i&zero(mes_i)&zero(dia_i) > ano_f&zero(mes_f)&zero(dia_f) Then '1 - Maior
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
				nome_unidade = " AND A053_desung like '%"&nome_unidade&"%'"
				campos = campos&",A053_desung, nm_sigla"
					cd_unidade = 0
					'response.write("não_numérico")
					if tipo_relatorio = "2" then
						nome_unidade replace(nome_unidade,"A053_desung","nm_unidade")
						campos = replace(nome_unidade,"A053_desung","nm_unidade")
					end if
			else
				nome_unidade = " AND a053_codung like '"&nome_unidade&"'"
				campos = campos&",A053_desung, nm_sigla"
					cd_unidade = 0
					'response.write("numérico")
					if tipo_relatorio = "2" then
						nome_unidade replace(nome_unidade,"A053_codung","cd_unidade")
						'campos = replace(nome_unidade,"A053_codung","cd_unidade")
					end if
			end if
			
	elseif cd_unidade > 0 then
		unidade = "AND a053_codung = "&cd_unidade&""
		
		campos = campos&",nm_sigla"
			if tipo_relatorio = "2" then
				unidade = replace(unidade,"a053_codung","cd_unidade")
			end if
	end if

'-------------------- Protocolo --------------------
	if nm_protocolo <> "" then
		protocolo = "AND a002_numpro like '%"&nm_protocolo&"%'"
		campos = campos&",a002_numpro"
	end if

'-------------------- Paciente --------------------
	if nm_paciente <> "" then
		if not IsNumeric(nm_paciente) then
			paciente = "AND a002_pacnom like '"&nm_paciente&"%'"
			campos = campos&",a002_pacreg,a002_pacnom"
		else
			paciente = "AND a002_pacreg like '"&nm_paciente&"%'"
			campos = campos&",a002_pacreg,a002_pacnom"
		end if
	end if

'-------------------- Idade --------------------
	if a002_pacida_i <> "" AND a002_pacida_f <> "" Then
		idade = "AND a002_pacida BETWEEN "&a002_pacida_i&" AND "&a002_pacida_f&""
	elseif a002_pacida_i <> "" AND a002_pacida_f = "" Then
		idade = "AND a002_pacida BETWEEN "&a002_pacida_i&" AND "&a002_pacida_i&""
	end if

'-------------------- Sexo --------------------
	if cd_sexo <> "" then
		sexo = "AND a002_pacsex='"&cd_sexo&"'"
	end if

'-------------------- Convenio --------------------
	if nm_convenio <> "" Then
		if not IsNumeric(nm_convenio) then
			convenio = "AND nm_convenio LIKE '%"&nm_convenio&"%'"
			campos = campos&",nm_convenio"
		else
			convenio = "AND nm_convenio = "&nm_convenio&""
			campos = campos&", nm_convenio"
		End if
	end if

'-------------------- Médico --------------------
	'Verifica os cirurgiões solicitados e monta o comando para o SQL
		
	if nm_medico <> "" then
		if not IsNumeric(nm_medico) then
			medico = "AND a055_desmed like '%"&nm_medico&"%'"
			campos = campos&", a055_desmed, a055_crmmed"
		else
			'medico = "AND cd_crm = "&nm_medico&""
			medico = "AND a055_crmmed = "&nm_medico&""
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
		'agenda_i = hora_i&":"&minuto_i
		'agenda_f = hora_f&":"&minuto_f
		'***
		
		'if campo_10_1 > 0 AND hora_i <> "" then
		'	If zero(hora_i)&zero(minuto_i) <= zero(hora_f)&zero(minuto_f) Then '1 - Menor ou igual
		'		agenda = "AND ({ fn hour(A002_horage) }) BETWEEN '"&hora_i&"' AND  '"&hora_f&"' AND ({ fn minute(A002_horage) }) BETWEEN '"&minuto_i&"' AND '"&minuto_f&"'" 
		'	elseIf zero(hora_i)&zero(minuto_i) > zero(hora_f)&zero(minuto_f) Then '1 - Maior
		'		ok = ""
		'		agenda_erro = "<b>O período selecionado é inválido!</b> O horário inicial é maior que o final.<br>&nbsp;"
		'	end if
		'End if

'-------------------- Hora Agenda --------------------

'-------------------- Rack --------------------
		if campo_11_1 > 0 Then
			if cd_rack <> "" Then
				rack = "AND a056_codrac = '"&cd_rack&"'"
			else
				rack = ""
			end if
			campos = campos&", a056_desrac"
		end if
		
		'if campo_11_2 > 0 Then
		'	if tipo_rack <> "" Then
		'		tiporack = "AND tipo_rack = '"&tipo_rack&"'"
		'	else
		'		tipo_rack = ""
		'	end if
		'	campos = campos&", tipo_rack"
		'end if
		
		
		if campo_11_2 > 0 then
			if cd_tipo_rack <> "" Then
				tipo_rack = "AND tipo_rack = "&cd_tipo_rack&""
			end if
		end if
		
'-------------------- Especialidade --------------------
		if campo_12_1 > 0 Then
			if cd_especialidade > "0" Then
				especialidade = "AND a057_codesp = '"&cd_especialidade&"'"
			end if
			'campos = campos&", a057_desesp"
			campos = campos&", nm_especialidade"
		end if

'-------------------- Procedimentos --------------------
		if campo_14_1 > 0 and str_procedimento <> "" AND tipo_relatorio = 2 Then
		'if campo_14_1 > 0 and tipo_relatorio = 2 Then
		
			if not IsNumeric(str_procedimento) then
				procedimento = "AND nm_procedimento like '%"&str_procedimento&"%'"
				campos = campos&", a058_codpro, nm_procedimento"
				'campos = campos&", cd_codigo, nm_procedimento"
			else
				procedimento = "AND a058_codpro = "&str_procedimento&""
				campos = campos&", cd_amb, nm_procedimento "
			end if
		end if

'-------------------- Materiais ------------------------
		if campo_15_1 > 0 and str_material <> "" AND tipo_relatorio = 3 Then
		'if campo_15_1 > 0 and tipo_relatorio = 3 Then
		
			if not IsNumeric(str_material) then
				material = "AND a061_nommat like '%"&str_material&"%'"
				campos = campos&", a061_codmat, a061_nommat"
			else
				material = "AND a061_codmat = "&str_material&""
				campos = campos&", a061_nommat "
			end if
		end if

'-------------------- Patrimonios ------------------------
		if campo_16_1 > 0 and str_patrimonio <> "" AND tipo_relatorio = 4 Then
		'if campo_16_1 > 0 and tipo_relatorio = 4 Then
		
			if not IsNumeric(str_patrimonio) then
				patrimonio = "AND nm_equipamento like '%"&str_patrimonio&"%'"
				campos = campos&", nm_tipo, cd_patrimonio, nm_equipamento"
			else
				patrimonio = "AND cd_patrimonio = "&str_patrimonio&""
				campos = campos&", nm_tipo, cd_patrimonio, nm_equipamento"
			end if
		end if

'**********************************************
'************** Seleção ***********************
'**********************************************
'selecao = " COUNT (a002_numpro) AS conta"
if selecao_top <> "" Then
	selecao_TOP = "TOP "&selecao_top
end if

'-------------------------------------------------------------
'--- Verifica qual o tipo de informaçao está sendo buscada ---
'-------------------------------------------------------------

'if str_procedimento <> "" AND str_material = "" AND str_patrimonio = "" AND tipo_relatorio = 2 Then
if tipo_relatorio = 2 Then
	objetivo = "a058_codpro"
	str_tabela = "VIEW_procedimento_protocolo_lista"
	codigo_indicador = ", cd_protocolo"
'elseif str_procedimento = "" AND str_material <> "" AND str_patrimonio = "" AND tipo_relatorio = 3 Then
elseif tipo_relatorio = 3 Then
	objetivo = "a061_codmat"
	str_tabela = "VIEW_material_protocolo_lista"
	codigo_indicador = ", cd_protocolo"
'elseif str_procedimento = "" AND str_material = "" AND str_patrimonio <> "" AND tipo_relatorio = 4 Then
elseif tipo_relatorio = 4 Then
	objetivo = "cd_codigo"
	str_tabela = "VIEW_patrimonio_protocolo_lista"
	codigo_indicador = ", cd_protocolo"
else
	objetivo = "a002_numpro"
	str_tabela = "VIEW_protocolo_lista"
		if campo_2_1 > 0 then
			codigo_indicador = ", cd_protocolo"
		end if
	
end if

selecao = selecao_top&" COUNT ("&objetivo&") AS conta "&codigo_indicador

if campos = "" Then
selecao = selecao
Else
selecao = selecao &", "&campos 
End if
	'Elimina uma possível duplicação de virgulas
	selecao = replace(selecao,", ,",",")

'**********************************************
'************** Condições *********************
condicao = "Where "&periodo&" "&unidade&" "&nome_unidade&" "&protocolo&" "&paciente&" "&idade&" "&sexo&" "&convenio&" "&medico&" "&agenda&" "&tipoagenda&" "&rack&" "&tipo_rack&" "&especialidade&" "&procedimento&" "&material&" "&patrimonio&"" 
'**********************************************
'************** Agrupamento *********************
if campos <> "" Then
agrupamento = "Group by "&campos&""&codigo_indicador
	teste_agrupamento = Instr(1,agrupamento,"Group by ,",0)
			if teste_agrupamento = 1 then
			agrupamento = replace(agrupamento,"Group by ,","Group by ")'Retira a virgula
			end if
			agrupamento = replace(agrupamento,", ,",",  ")'Retira a virgula dupla
			
			'response.write(teste_agrupamento)

'**********************************************
'**************** Ordem ***********************
sentido = request("sentido")

ordem_inicial = request("ordem_inicial")
		ordem_1 = request("ordem_1")
		'ordem_2 = request("ordem_2")
		ordem_3 = request("ordem_3")
		ordem_4 = request("ordem_4")
		ordem_5 = request("ordem_5")
		'ordem_6 = request("ordem_6")
		'ordem_7 = request("ordem_7")
		'ordem_8 = request("ordem_8")
		ordem_9 = request("ordem_9")
		'ordem_10 = request("ordem_10")
		'ordem_11 = request("ordem_11")
		ordem_12 = request("ordem_12")
		ordem_13 = request("ordem_13")
		ordem_14 = request("ordem_14")
		ordem_15 = request("ordem_15")
		'ordem_16 = request("ordem_16")
		'ordem_17 = request("ordem_17")
%>

<!--include file="../includes/ordem_protocolos.asp"-->

<%ordem_total = request("ordem_total")
	if campo_9_0 > 0 then
		ordem_total = replace(ordem_total,"a002_datpro","dt_ano")
		ordem_total = replace(ordem_total,"dt_ano,dt_mes","dt_ano")
	elseif campo_9_1 > 0 then
		if ano_i = ano_f then 
			str_data = "dt_mes" 
		elseif ano_i<>ano_f then 
			str_data = "dt_ano,dt_mes" 
		end if
		ordem_total = replace(ordem_total,"a002_datpro",str_data)
		'ordem_total = replace(ordem_total,"a002_datpro",str_data)
		
	elseif campo_9_2 > 0 then
		if ano_i = ano_f then 
			str_data = "dt_mes" 
		elseif ano_i<>ano_f then 
			str_data = "dt_ano,dt_mes" 
		end if
		ordem_total = ordem_total
	end if
	
	ver_string_ordem = instr(ordem_total,",")
	
		if ver_string_ordem = 1 then
			ordem_total = mid(ordem_total,2,len(ordem_toral))
		end if
	
			if ordem_total <> "" Then
				ordem = "ORDER BY "&ordem_total&" "&sentido 
			end if
			
				ordem_primaria = instr(ordem_total,",")
					if ordem_primaria = 0 then
						ordem_primaria = len(ordem_total)
					end if
						primeiro_campo = trim(mid(ordem_total,1,ordem_primaria))

end if%>
<head>
</head>
<body>
			
			<table border="1" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" class="no_print" width="500">
				<form action="protocolo.asp" method="post" name="form">
				<input type="hidden" name="tipo" value="med_cirurgias">
				<!-- 3ª Etapa -->
				<input type="hidden" name="ordem_res">
				<input type="hidden" name="ordem_total" value="<%=ordem_total%>">
				<input type="hidden" name="ordem_inicial" value="<%if ordem_inicial = "" Then%>1<%else response.write(ordem_inicial) end if%>">
				<tr class="no_print">
					<td>&nbsp;</td>
					<td bgcolor="#C0C0C0" align="center" colspan="100%"><b>BUSCA AVANÇADA - <font color="red">Cirurgias por Medicos</font></b></td>
				</tr>
				<tr class="no_print">
					<td colspan="1"><img src="imagens/px.gif" width="10" height="1"></td>
					<td colspan="1"><img src="imagens/px.gif" width="10" height="1"></td>
					<td colspan="3"><img src="imagens/px.gif" width="110" height="1"></td>
					<td colspan="4"><img src="imagens/px.gif" width="525" height="1"></td>
					<td><img src="imagens/px.gif" width="10" height="1"></td>
				</tr>
				<tr><td colspan="100%" align="center"><%=compara%></td></tr>
				<tr class="no_print" bgcolor="#e6e6e6">
					<td rowspan="22" bgcolor="#ffffff">&nbsp;</td>
					<td colspan="1">&nbsp;<b>TIPO</b></td>
					<td colspan="3">&nbsp;<b>MOSTRAR</b></td>
					<td colspan="4">&nbsp;<b>DETALHES</b></td>
					<td>&nbsp;<b>ORDEM</b></td>
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr>
					<td>&nbsp;</td>
					<td>&nbsp;<a href="javascript:selecionar_tudo()">tudo</a> :: <a href="javascript:deselecionar_tudo()">nada</a></td>
					<td colspan="6">&nbsp;</td>
					<td>&nbsp;<a href="javascript:void();" onClick="limpa_ordem()">Limpa Ordem</a></td>
				</tr>
<!-- *** Data *** -->
				<tr bgcolor="#eaeaea" class="no_print" onmouseover="mOvr(this,'#ccffff');" onclick="javascript:;" onmouseout="mOut(this,'#eaeaea');">
					<td>&nbsp;</td>					
					<td colspan="3">
						<%if data <> 0  then%>
					<%ck_data = "checked"%>
					<%end if%>						
						<input type="checkbox" name="data" value="1" <%=ck_data%>>&nbsp;Data</td>
					<td colspan="4" valign="middle">
					<%if campo_9_0 > 0 then%>
						<%ck_campo_9_0 ="checked"%>
					<%elseif campo_9_1 > 0 then%>
						<%ck_campo_9_1 ="checked"%>
					<%else'if campo_9_2 > 0 then%>
						<%ck_campo_9_2 ="checked"%>
					<%end if%>
						<input type="radio" name="intervalo" value="so_ano" <%=ck_campo_9_0%>>Anual
						<input type="radio" name="intervalo" value="dt_ano, dt_mes" <%=ck_campo_9_1%>>mensal
						<input type="radio" name="intervalo" value="a002_datpro" <%=ck_campo_9_2%>>diário
						&nbsp;&nbsp;&nbsp;
						<select name="dia_i" class="inputs">
							<%numero = 1
							do while NOT numero = 32
							if int(dia_i) = numero Then
							check = "selected"
							end if%>					
							<option value="<%=numero%>" <%=check%>><%=zero(numero)%></option>
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
					<td>
					<%ordem_var = ordem_1
					ordem_num="ordem_1"
					ordem_campo="a002_datpro"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>
<!-- *** Quantidade *** -->
				
<!-- *** Unidade *** -->
				<tr bgcolor="#eaeaea" class="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'#eaeaea');">
					<td>&nbsp;</td>
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
						<option value="<%=cd_uni%>" <%=ck_uni%>><%=nm_sigla%> - <%=nome_unidade%> - (cód: <%=zero(cd_uni)%>)</option>
							<%rs_uni.movenext
							ck_uni = ""
							wend
							%>
					</select>
					OU
					<input type="text" name="nm_unidade" value="<%=nm_unidade%>" size="20" maxlength="200"  class="inputs">
					
					</td>
					<td><%ordem_var = ordem_3
					ordem_num="ordem_3"
					ordem_campo="nm_sigla"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>
<!-- *** Nº Protocolo *** -->
				<tr bgcolor="#f7f7f7" class="no_print" onmouseover="mOvr(this,'#ccffff');" onclick="javascript:;" onmouseout="mOut(this,'');">
					<td>&nbsp;<input type="radio" name="tipo_relatorio" value="1" <%if tipo_relatorio <= 1 then response.write("checked") end if%>></td>
					<td colspan="3"><%if campo_2_1 <> 0 then%><%ck_campo_2 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_numpro" <%=ck_campo_2%>>&nbsp;Protocolo						
					</td>
					<td colspan="4"><!--input type="text" name="nm_und" size="2" maxlength="3" class="inputs" value="<%'=nm_und%>" class="inputs"--><input type="text" name="nm_protocolo" size="10" maxlength="50" class="inputs" value="<%=mm_protocolo%>" class="inputs">
						
					</td>
					<td><%ordem_var = ordem_4
					ordem_num="ordem_4"
					ordem_campo="a002_numpro"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>
<!-- *** Paciente *** -->
				<tr bgcolor="#eaeaea" class="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'#eaeaea');">
					<td>&nbsp;</td>
					<td colspan="3"><%if campo_3_1 <> 0 OR nm_paciente <> "" then%><%ck_campo_3 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a002_pacreg, a002_pacnom" <%=ck_campo_3%>>&nbsp;Paciente						
					</td>
					<td colspan="4"><input type="text" name="nm_paciente" size="50" maxlength="500" value="<%=nm_paciente%>" class="inputs"></td>
					<td><%ordem_var = ordem_5
					ordem_num="ordem_5"
					ordem_campo="a002_pacnom"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>
			
<!-- *** Médico *** -->
				<tr bgcolor="#eaeaea" class="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'eaeaea');">
					<td>&nbsp;</td>
					<td colspan="3"><%if campo_7_1 <> 0 OR a055_desmed <> "" then%><%ck_campo_7 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value=",a055_crmmed, a055_desmed" <%=ck_campo_7%>>&nbsp;Médico
					</td>
					<td colspan="4"><input type="text" name="nm_medico" size="50" maxlength="500" value="<%=nm_medico%>" class="inputs">
						
					</td>
					<td><%ordem_var = ordem_9
					ordem_num="ordem_9"
					ordem_campo="a055_desmed"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>

<!-- *** Tipo de agenda *** -->
				<tr bgcolor="#eaeaea" class="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'eaeaea');">
					<td>&nbsp;</td>
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
					<td><%ordem_var = ordem_12
					ordem_num="ordem_12"
					ordem_campo="a002_tipage"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>

<!-- *** Rack *** -->
				<tr bgcolor="#f7f7f7" id="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'');">
					<td>&nbsp;</td>
					<td colspan="3"><%if campo_11_1 <> 0 then%><%ck_campo_11 = "checked"%><%end if%>
						<!--input type="checkbox" name="campos" value=",a056_codrac, a002_numseq" <%=ck_campo_11%>>&nbsp;Rack-->
						<input type="checkbox" name="campos" value=",a056_codrac, cd_protocolo" <%=ck_campo_11%>>&nbsp;Rack
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
					<td><%ordem_var = ordem_13
					ordem_num="ordem_13"
					ordem_campo="a056_codrac"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>
<!-- *** Rack tipo*** -->
				<tr bgcolor="#f7f7f7" id="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'');">
					<td>&nbsp;</td>
					<td colspan="3"><%if campo_11_2 <> 0 then%><%ck_campo_11_2 = "checked"%><%end if%>
						<input type="checkbox" name="campos" value="tipo_rack" <%=ck_campo_11_2%>>&nbsp;Rack Tipo
					</td>
					<td colspan="4">
						<%if cd_tipo_rack = "1" Then
							ck_racktipo1 = "selected"
						elseif cd_tipo_rack = "2" Then
							ck_racktipo2 = "selected"
						end if%>
						<select name="cd_tipo_rack" class="inputs"> 
							<option value="">&nbsp;</option>
							<option value="1" <%=ck_racktipo1%>>Rack Vdlap</option>
							<option value="2" <%=ck_racktipo2%>>Rack Cirurgião</option>						
						</select>&nbsp;
					</td>
					<td><%ordem_var = ordem_14
					ordem_num="ordem_14"
					ordem_campo="tipo_rack"%>
					<img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>	
<!-- *** Procedimento *** -->
				<tr bgcolor="#f7f7f7" class="no_print" onmouseover="mOvr(this,'#ccffff');"  onclick="javascript:;" onmouseout="mOut(this,'');">
					<td rowspan="1">&nbsp;<!--input type="radio" name="tipo_relatorio" value="2" <%if tipo_relatorio = 2 then response.write("checked") end if%>--></td>
					<td colspan="3"><%if campo_14_1 <> 0 then%><%ck_campo_14 = "checked"%><%end if%>
						<!--input type="checkbox" name="campos" value=",a058_codpro" <%=ck_campo_14%>>&nbsp;Procedimento-->
						<input type="checkbox" name="procedimentos" value="procedimentos,a002_numpro" <%=ck_campo_14%>>&nbsp;Procedimento
					</td>
					<td colspan="4"><input type="text" name="nm_procedimento" size="50" maxlength="50" class="inputs" value="<%=str_procedimento%>">&nbsp;</td>
					<td><%ordem_var = ordem_15
					ordem_num="ordem_15"
					ordem_campo="a058_despro"%>
					<!--img src="../../imagens/blackdot.gif" alt="Inclui" width="15" height="15" border="0" onClick="ordem('<%=ordem_campo%>',document.form.ordem_total.value,'<%=ordem_num%>')" id="<%=ordem_num%>" name="<%=ordem_num%>" <%if ordem_var <> "" Then%>style="display:none;"<%end if%>--><input type="text" name="<%=ordem_num%>" size="2" maxlength="2" readonly="1" style="border: 1px solid none;" value="<%=ordem_var%>"></td>
				</tr>

<%'end if%>
<!-- *** Subtotal-2 ***-->
				<tr bgcolor="#eaeaea"><%if subtotais > 0 then%><%ck_subtotais = "checked"%><%end if%>
					<td>&nbsp;</td>
					<td colspan="3"><input type="checkbox" name="subtotais" value="1" <%=ck_subtotais%>>&nbsp;Subtotais.</td>
					<td colspan="6" style="face_size:7px; color:silver;" valign="baseline">&nbsp; Obs.: Para obter os subtotais desejados, ordene a busca pela coluna à direita.</td>
					
				</tr>
				<!--tr>
					<td colspan="100%">.</td>
				</tr-->
<!-- ***********************************************************************************************-->
<!-- *********************************      ORDEM      *********************************************-->
<!-- ***********************************************************************************************-->

				<tr class="no_print">
					<td colspan="8">&nbsp;<font color="#ff0000"><%=alerta%> <%'=dia_f%></font></td>
					<td>
						<select name="sentido" size="1">
							<%if sentido = "DESC" Then%>
								<%cres="selected"%>
							<%elseif sentido = "DESC" Then%>
								<%decr="selected"%>
							<%end if%>
							<option value=""></option>
							<option value="ASC" <%=cres%>>Cres.</option>
							<option value="DESC" <%=decr%>>Decres.</option>
						</select>
					</td>
				</tr>		
				<tr><td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<tr class="no_print">
					<td colspan="8" align="center">
						<!--input type="hidden" name="x" value="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->
						<input type="submit" name="ok" value="Buscar">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input onclick="javascript:JsRedefinir('')" type="button" value="Limpar" title="Redefinir"> 
					</td>
				</tr>
				<tr>
					<td colspan="100%"><img src="../imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<tr class="txt">
					<td align="left" colspan="100%">
					<font color="red"><%=periodo_erro%></font><br>
					</td>
				</tr>
			<%if session_cd_usuario = 46 then%>
				<tr class="no_print">
					<td align="left" colspan="100%" bgcolor="#ffff00">
					<b>SELECT </b><%=selecao%><br>
					<b>FROM <%=str_tabela%> </b><br>
					<%=condicao%> <%=cond_periodo%> <%=cond_und%> <br>
					<%=agrupamento%><br>
					<%=ordem%> &nbsp;<%'=sentido%><br>
					</td>
				</tr>
				<%end if%>
			</table>
			<br class="no_print"><br class="no_print">
			
			<blockquote >
			<table align="Left" border="0" cellpadding="1" cellspacing="1" width="10">
			
			</form>
<%
'ok = ""
	cabecalho = 1
	if ok <> "" Then 'Inicio do resultado da pesquisa
			'******************************
			'*							  *
			'* 	  2ª Parte - Resultados   *
			'*                            *
			'******************************	%>
				
								
					
				
<%

xsql = "SELECT "&selecao&" FROM "&str_tabela&" "&condicao&" "&agrupamento&" "&ordem&""' "&sentido&""
Set rs_pc = dbconn.execute(xsql)
linhas = 0
do while Not rs_pc.EOF
	linhas = 1
rs_pc.movenext
pre_conta = linhas + pre_conta
loop%>
<%if pre_conta > 700 Then%>
	<tr><td colspan="100%">A busca retornou muitas linhas, por favor refine a sua busca, para . (<%=linhas%>/<%=pre_conta%>)</td></tr>			
<%else%>
				<!--tr><td colspan="100%" class="ok_print"><img src="imagens/blackdot.gif" alt="" width="100%" height="1">x</td></tr-->
	<%	
			principal = primeiro
			nr = 0
			header = 1
			n_linha = 1
			n_cor = 1
			
			xsql = "SELECT "&selecao&" FROM "&str_tabela&" "&condicao&" "&agrupamento&" "&ordem&"  "
			Set rs = dbconn.execute(xsql)
			
			do while Not rs.EOF
			
			conta = rs("conta")
			
			if n_cor = "1" then
				l_cor = "#dddddd"
			else
				l_cor = "#f5f5f5"
			end if
			
			
			if campo_1_1 > 0 OR cd_unidade > 0 OR nm_unidade <> "" then
				if tipo_relatorio = 2 then
					cd_unidade = rs("cd_unidade")
				else
					cd_unidade = rs("a053_codung")
				end if
					'cd_unidade = rs("a053_codung")
					a053_codung = cd_unidade&"."
				nm_und = rs("nm_sigla")	
					if primeiro_campo = "nm_sigla" Then
						compara = cd_unidade
					end if
				
				strsql_unid = "SELECT * FROM TBL_unidades where cd_codigo = "&a053_codung&""
				Set rs_unid2 = dbconn.execute(strsql_unid)
					while not rs_unid2.EOF
						nm_unidade = rs_unid2("nm_unidade")
					rs_unid2.movenext
					wend
							
			end if
			
			if campo_2_1 > 0 then
				a002_numpro = rs("a002_numpro")
					if primeiro_campo = "a002_numpro" Then
						compara = a002_numpro
					end if
			end if
			
			if campo_3_1 > 0 OR nm_paciente <> "" then
				a002_pacreg = rs("a002_pacreg")
				a002_pacnom = rs("a002_pacnom")
					if primeiro_campo = "a002_pacnom" Then
						compara = a002_pacnom
					end if
			end if
			
			if campo_4_1 > 0 then
				a002_pacida = rs("a002_pacida")
					if primeiro_campo = "a002_pacida" Then
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
					if primeiro_campo = "a002_pacsex" Then
						compara = a002_pacsex
					end if
			end if
			
			if campo_6_1 > 0 then
				nm_convenio = rs("nm_convenio")
					if primeiro_campo = "nm_convenio" Then
						compara = nm_convenio
					end if
			end if
			
			if campo_7_1 > 0 then
				cirurgiao = rs("a055_desmed")
				crm = rs("a055_crmmed")
					if primeiro_campo = "a055_desmed" Then
						compara = cirurgiao
					end if
			end if
			
			'Não há campo_8 ainda
			
			'****************************************************
			'*** Data - tratamento da data em relação à ordem ***
			'****************************************************
			If data = 1 AND campo_9_0 > 0 then
				dt_ano = rs("dt_ano")
				A002_datpro = dt_ano
					if ordem_data = 1 Then
					compara = 1'A002_datpro
					end if
			Elseif data = 1 AND campo_9_1 > 0 then
				dt_mes = rs("dt_mes")
				dt_ano = rs("dt_ano")
				
				if ano_i = ano_f then
				
				end if
				
				A002_datpro = left(mes_selecionado(dt_mes),3)&"/"&dt_ano
					if primeiro_campo = "dt_ano, dt_mes" Then
						compara = A002_datpro
					elseif primeiro_campo = "dt_ano" Then
						compara = year(A002_datpro)
					elseif primeiro_campo = "dt_mes" Then
						compara = month(A002_datpro)
					end if
			Elseif data = 1 AND campo_9_2 > 0 then
				A002_datpro_dia = day(rs("A002_datpro"))
				A002_datpro_mes = month(rs("A002_datpro"))
				A002_datpro_ano = year(rs("A002_datpro"))
					A002_datpro = zero(A002_datpro_dia)&"/"&zero(A002_datpro_mes)&"/"&A002_datpro_ano			
						if primeiro_campo = "a002_datpro" Then
							compara = a002_datpro
						end if
			end if
			'*******************************************************
			
			if campo_10_1 > 0 then
				agenda_hr = zero(hour(rs("a002_horage")))&":"&zero(minute(rs("a002_horage")))
					if primeiro_campo = "a002_horage" Then
						compara = agenda_hr
					end if
			end if
			
			if campo_10_2 > 0 then
				cirurgia_hr_ini = zero(hour(rs("a002_horini")))&":"&zero(minute(rs("a002_horini")))
				cirurgia_hr_ter = zero(hour(rs("a002_horfin")))&":"&zero(minute(rs("a002_horfin")))
					if primeiro_campo = "a002_horini" Then
						compara = cirurgia_hr_ini
					end if
			end if
			
			if campo_11_1 > 0 then
				a056_codrac = rs("a056_codrac")
				a056_desrac = rs("a056_desrac")
					if primeiro_campo = "a056_codrac" Then
						compara = a056_desrac
					end if
			end if
			
			if campo_11_2 > 0 then
				'a056_codrac = rs("a056_codrac")
				tipo_rack = rs("tipo_rack")
					if tipo_rack = 1 then
						nm_tipo_rack = "Vd Lap"
					elseif tipo_rack = 2 then
						nm_tipo_rack = "Cirurgião"
					else
						nm_tipo_rack = "Ñ espec."
					end if 
					if primeiro_campo = "tipo_rack" Then
						compara = tipo_rack
					end if
			end if
			
			if campo_12_1 > 0 then
				a057_codesp = rs("a057_codesp")
				nm_especialidade = rs("nm_especialidade")
					if primeiro_campo = "a057_codesp" Then
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
					if primeiro_campo = "a002_tipage" Then
						compara = a002_tipage
					end if
			end if
			
			if campo_14_1 > 0 then
				'cd_protocolo = rs("cd_protocolo")
			end if
			
			if campo_15_1 > 0 then
				'cd_protocolo = rs("cd_protocolo")
			end if
			
			if campo_16_1 > 0 then
				'cd_protocolo = rs("cd_protocolo")
			end if
			
			
			%>
			
			
			<%if cabecalho <> "" then
				if cabecalho = 1 then
					conta_cabecalho = conta_cabecalho + 1%>
					<tr>					
						<td bgcolor="#C0C0C0" align="center" colspan="100%">  <b>RELATÓRIO DE PROCEDIMENTOS MÉDICOS - Pág.: <%=zero(conta_cabecalho)%></b></td>
					</tr>
					<tr>					
						<td bgcolor="#C0C0C0" align="center" colspan="100%">  <b><%=nm_unidade%> - <%=mes_selecionado(mes_i)&"/"&ano_i%> a <%=mes_selecionado(mes_f)&"/"&ano_f%></b></td>
					</tr>
					<tr>					
						<td bgcolor="#C0C0C0" align="center" colspan="100%">  <b>Dr.: <%=cirurgiao%></b></td>
					</tr>
				<%cabecalho = 2
				end if%>
				<%if cabecalho = 2 then%>
				<tr bgcolor="#c0c0c0">
						<td>Nº<br><img src="../imagens/px.gif" alt="" width="15" height="1" border="0"></td>
						<!--td>Qtd<br><img src="../imagens/px.gif" alt="" width="22" height="1" border="0"></td-->
						<%'if campo_1_1 > 0 OR cd_unidade > 0 OR nm_unidade <> "" then%><!--td>Unid.<br><img src="../imagens/px.gif" alt="" width="40" height="1" border="0"></td--><%'end if%>
						<%if campo_2_1 > 0 then%><td>Protocolo<br><img src="../imagens/px.gif" alt="" width="75" height="1" border="0"></td><%end if%>
						<%if campo_3_1 > 0 OR nm_paciente <> "" then%><td align="left" colspan="2">Paciente<br><img src="../imagens/px.gif" alt="" width="220" height="1" border="0"></td><%end if%>
						<%if campo_4_1 > 0 then%><td>Idade<br><img src="../imagens/px.gif" alt="" width="35" height="1" border="0"></td><%end if%>
						<%if campo_5_1 > 0 then%><td>Sexo<br><img src="../imagens/px.gif" alt="" width="30" height="1" border="0"></td><%end if%>
						<%if campo_6_1 > 0 then%><td>Convênio<br><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td><%end if%>
						<%'if campo_7_1 > 0 then%><!--td colspan="2">Equipe Médica<br><img src="../imagens/px.gif" alt="" width="270" height="1" border="0"></td--><%'end if%>
						<%if data > 0 then%><td>Data<br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td><%end if%>
						<%if campo_10_1 > 0 then%><td>Agenda<br><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td><%end if%>
						<%if campo_10_2 > 0 then%><td>Inicio<br><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
												  <td>Fim<br><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
												  <td>Tempo<br><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td>
												  <td>Atrasos<br><img src="../imagens/px.gif" alt="" width="45" height="1" border="0"></td><%end if%>
						<%if campo_11_1 > 0 then%><td>Rack<br><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td><%end if%>
						<%if campo_11_2 > 0 then%><td>Rack tipo<br><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td><%end if%>
						<%if campo_12_1 > 0 then%><td>Especialidade<br><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td><%end if%>
						<%if campo_13_1 > 0 then%><td>tipo<br><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td><%end if%>
						<%if campo_14_1 > 0 then%><td>Procedimento<br><img src="../imagens/px.gif" alt="" width="250" height="1" border="0"></td><%end if%>
						<%if campo_15_1 > 0 then%><td>Materiais<br><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td><%end if%>
						<%if campo_16_1 > 0 then%><td>Patrimonio<br><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td><%end if%>
						<%'if campo_2_1 > 0 then%><!--td class="no_print"></td--><%'end if%>
					</tr>
				<%end if%>
			<%cabecalho = 3
			end if%>
			
			
			<%if ordem_unidade = 1 and header = 1 then%>
				<tr><td colspan="100%" align="center"><b><%=nm_und%></b></td></tr>
			<%end if%>	
				
				<%if grava_info <> compara AND nr > 0 AND subtotais > 0 Then
				'if  grava_info <> compara AND subtotais > 0 Then%>
				<%conta_linha = conta_linha + 1%>
				<tr><td bgcolor="#ebebeb" colspan="100%"><b><i><%=sub_conta%></i></b></td></tr><tr><td>&nbsp;</td></tr>
				<%sub_conta = 0
				end if%>
				<tr bgcolor="<%=l_cor%>" onmouseover="mOvr(this,'#CFC8FF');" onclick="javascript:;" onmouseout="mOut(this,'<%=l_cor%>');" >				
				<%conta_linha = conta_linha + 1%>
					<!--td><%'=primeiro_campo&" <br> "&grava_info&" <br> "&compara&" <br> "&nr&" <br> "&subtotais%></td-->
					<td valign="top" style="color:gray;"<%if conta_linha = "40" then%>style="page-break-after:always;"<%end if%>><%=zero(n_linha)%><br><%'=conta_linha%></td>
					<!--td valign="top"><%'=zero(conta)%></td-->
				<%'if campo_1_1 > 0 OR cd_unidade > 0 OR nm_unidade <> "" AND header <> 1 then%>
					<!--td valign="top"><%=nm_und%></td--><%'end if%>
				<%if campo_2_1 > 0 then%>	
					<td valign="top"><%if cd_unidade <> "" AND a002_numpro <> "" Then%>
					<a href="javascript: void(0);" onclick="window.open('protocolo/protocolo_folha.asp?tipo=ver&cd_unidade=<%=cd_unidade%>&cd_protocolo=<%=proto(a002_numpro)%>&cd_digito=<%=right(digitov(cdun(a053_codung)&proto(a002_numpro)),2)%>&janelas_cadastros', '', 'scrollbars=1, width=461, height=570' ); return false;">
					<%end if%>
					<%'=a053_codung&(a002_numpro)%>
					<%'=a053_codung&"a"&a002_numpro%>
					<%if tipo_relatorio = 2 then%>				
						<%=digitov(cdun(cd_unidade)&proto(no_protocolo))%>
					<%else%>
						<%=digitov(cdun(a053_codung)&proto(a002_numpro))%>
					<%end if%>
					<!--input type="checkbox" name="nada" value=""--></a></td><%end if%>			
				<%if campo_3_1 > 0 OR nm_paciente <> "" then%>
					<td valign="top"><%=a002_pacreg%></td><td valign="top"><%=a002_pacnom%></td><%end if%>
				<%if campo_4_1 > 0 then%>
					<td valign="top"><%=a002_pacida%></td><%end if%>
				<%if campo_5_1 > 0 then%>
					<td valign="top"><%=sexo%></td><%end if%>
				<%if campo_6_1 > 0 then%>
					<td valign="top"><%=a054_codcon%> &nbsp; <%=nm_convenio%></td><%end if%>
				<%'if campo_7_1 > 0 then%>
					<!--td valign="top"><%=crm%></td><td valign="top">Dr. <%=cirurgiao%></td--><%'end if%>
				<%if data > 0 then%>
					<td valign="top"><%=zero(day(A002_datpro))&"/"&zero(month(A002_datpro))&"/"&right(year(A002_datpro),2)%> <%'=compara%><%'=grava_info%></td><%end if%>
				<%if campo_10_1 > 0 then%>
					<td valign="top" align="right"><i><%=agenda_hr%></i></td><%end if%>
				<%if campo_10_2 > 0 then%>
					<td valign="top" align="right"><%=cirurgia_hr_ini%></td>
					<td valign="top" align="right"><%=cirurgia_hr_ter%></td><%'end if%>
					<%'if campo_10_3 > 0 then%>
				
				<%'transforma horas em minutos:-------------------------------------------------------
				'if cirurgia_hr_ini = "" or cirurgia_hr_ter = "" then
				if cirurgia_hr_ini <> "" then
					hora_agenda = mid(agenda_hr,1,2)
					min_agenda = mid(agenda_hr,4,5)
						hragenda = hora_agenda * 60 + min_agenda
					
					hora_inicio = mid(cirurgia_hr_ini,1,2)
					min_inicio = mid(cirurgia_hr_ini,4,5)
						hrinicio = hora_inicio * 60 + min_inicio
					
					hora_final = mid(cirurgia_hr_ter,1,2)
					min_final = mid(cirurgia_hr_ter,4,5)
						hrfinal = hora_final * 60 + min_final
					
					'--------------------------------------------------------------------------------------
					'if asd = 1 then
						'*** Tempo decorrido da cirurgia ***
							if hrfinal > hrinicio then
								cirurgia_tempo = formatahora(hrfinal - hrinicio)
								'cirurgia_tempo = formatahora(hrfinal - hrinicio) 'mesmo dia
							else
								if hrfinal = "0" then
									'cirurgia_tempo = formatahora("1440" - hrinicio)
									cirurgia_tempo = formatahora("1440" - hrinicio) 'dia posterior
								else
									dia_inicio = ("1440" - hrinicio) 'dia posterior
									dia_final = hrfinal
									cirurgia_tempo = formatahora(dia_inicio + hrfinal) 
								end if
							end if
	
						'*** Atraso no inicio da cirurgia ***
							 if hrinicio > hragenda then
								'cirurgia_atraso = "-"&formatahora(hrinicio - hragenda) 
								cirurgia_atraso = "-"&formatahora(hrinicio - hragenda) 'ATRASADAS
							else
								'cirurgia_atraso = "&nbsp;"&formatahora(hragenda - hrinicio)
								cirurgia_atraso = "&nbsp;"&formatahora(hragenda - hrinicio) 'ANTECIPADAS
							end if
					end if
				%>
					
					<td valign="top" align="right"><b><%=cirurgia_tempo%></b></td>
					<td valign="top" align="right"><%=cirurgia_atraso%></td><%end if%>
					<%if campo_10_4 > 0 then%>
					<td valign="top" align="right"><%=cirurgia_tempo%></td><%end if%>
				<%if campo_11_1 > 0 then%>
					<td valign="top"><%=a056_desrac%></td><%end if%>
				<%if campo_11_2 > 0 then%>
					<td valign="top"><%=nm_tipo_rack%><%'=tipo_rack%></td><%end if%>
				<%if campo_12_1 > 0 then%>
					<td valign="top"><%=nm_especialidade%></td><%end if%>	
				<%if campo_13_1 > 0 then%>
					<td valign="top"><%=tipo_agenda%></td><%end if%>
				<%if campo_14_1 > 0 then%>	
					<td valign="top" width="800">						
						<%if campo_2_1 > 0 then
						cd_protocolo = rs("cd_protocolo")
							
							if tipo_relatorio = 2 then
								proc_filtro = " AND nm_procedimento like '%"&str_procedimento&"%'"
							end if
							
							strsql_proc = "SELECT * FROM View_protocolo_procedimento WHERE cd_protocolo = "&cd_protocolo&" "&proc_filtro&" ORDER BY a003_propri desc"
							Set rs_proc = dbconn.execute(strsql_proc)%>
								<%'while not rs_proc.EOF
								'a002_numpro = rs_proc("a002_numpro").
								cod_amb = rs_proc("a058_codpro")
								nm_procedimento = rs_proc("nm_procedimento")
									nm_proc_len = len(nm_procedimento)
										if nm_proc_len > 55 then
										nm_procedimento = left(nm_procedimento,55)&"..."
										end if
								a003_propri = rs_proc("a003_propri")
								'if a003_propri <> "S" then%>
									<!--i><font color="red"-->
								<%'end if%><!--b-->
								<%'=a058_codpro%><!--/b></font--><%'=lcase(a003_propri)%><%'if a003_propri = "s" then%><!--/i--><%'end if%>
								<%conta_tipo_2 = conta_tipo_2 + 1%>
								<%=mid(cod_amb,1,2)&"."&mid(cod_amb,3,2)&"."&mid(cod_amb,5,3)&"."&mid(cod_amb,8,1)%> - <%=nm_procedimento&crlf%><!--br-->							
								<%
								
								rs_proc.movenext
								'wend
								
								rs_proc.close
								set rs_proc = nothing%>
								
							<%Else%>
								<i style="color:red;">Atenção: O campo Protocolo  deve ser marcado.</b>
							<%end if%>
							
							
					 </td>
				<%end if%>
				<%if campo_15_1 > 0 then%>					
					<td valign="top" width="800">
						<%if campo_2_1 > 0 then
						cd_protocolo = rs("cd_protocolo")
							
							if tipo_relatorio = 3 then
								mat_filtro = " AND a061_nommat like '%"&str_material&"%'"
							end if
							
							strsql_mat = "SELECT * FROM view_protocolo_material WHERE cd_protocolo = "&cd_protocolo&" "&mat_filtro&""
							Set rs_mat = dbconn.execute(strsql_mat)%>
								<%while not rs_mat.EOF
								a061_nommat = rs_mat("a061_nommat")
								a010_quantidade = rs_mat("a010_quantidade")%>
								<%conta_tipo_3 = conta_tipo_3 + a010_quantidade%>
								<%=zero(a010_quantidade)&" - "&a061_nommat%><br>							
								<%rs_mat.movenext
								wend
								
								rs_mat.close
								set rs_mat = nothing%>
						<%Else%>
							<i style="color:red;">Atenção: O campo Protocolo  deve ser marcado.</b>
						<%end if%>			
					 </td>
				<%end if%>
				<%if campo_16_1 > 0 then%>
					<td valign="top" width="800">
						<%if campo_2_1 > 0 then
						cd_protocolo = rs("cd_protocolo")
						
							strsql_pat = "SELECT * FROM view_protocolo_patrimonio_lista WHERE cd_protocolo = '"&cd_protocolo&"'"
							Set rs_pat = dbconn.execute(strsql_pat)%>
								<%while not rs_pat.EOF
								nm_tipo = rs_pat("nm_tipo")
								cd_patrimonio = rs_pat("cd_patrimonio")
								nm_equipamento = rs_pat("nm_equipamento")%>
								<%conta_tipo_4 = conta_tipo_4 + 1%>
								<%=nm_tipo&cd_patrimonio&" - "&nm_equipamento%><br>							
								<%rs_pat.movenext
								wend
								
								rs_pat.close
								set rs_pat = nothing%>
						<%Else%>
							<i style="color:red;">Atenção: O campo Protocolo  deve ser marcado.</b>
						<%end if%>			
					 </td>
				<%end if%>
				
				
				
				
				<%'if campo_2_1 > 0 then%>	
					<!--td valign="top" class="no_print"><%if cd_unidade <> "" AND a002_numpro <> "" Then%><%'response.write("<a href=protocolo.asp?tipo=dvd&cd_unidade="&cd_unidade&"&cd_protocolo="&a002_numpro&">")%><%end if%>
					<img src="../imagens/filme3.gif" alt="" width="10" height="11" border="0"></a></td-->
				<%'end if%>	
				
				</tr>
				<%if conta_linha = 40 then%>
				<tr><td colspan="100%"></td></tr>
				
				<%conta_linha=0
				cabecalho = 1
				end if%>
				
				
			<%rs.movenext
				grava_info = compara
				sub_conta = sub_conta + conta
			geral = geral + conta
			nr = 1
			n_linha = n_linha + 1
			n_cor = n_cor + 1
				if n_cor > 2 Then
					n_cor = 1
				end if
			header = header + 1
			loop
			
			rs.close
			set rs = nothing
end if%>
			<%if subtotais = 1 then%>
				<tr><td bgcolor="#ebebeb" colspan="100%"><b><i><%=sub_conta%></i></b></td></tr><tr><td>&nbsp;</td></tr>
				<%end if%>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>					
				<tr bgcolor="#e1e1e1">
						<%if tipo_relatorio = 1 then%>	
							<td colspan="5">&nbsp;Total: <b><%=geral%></b><br><%=agenda_erro%></td>
							
							<%'if tipo_relatorio = 1 then%>								
								<td colspan="5">&nbsp;<%=agenda_erro%></td>
							<%elseif tipo_relatorio = 2 then%>
								<td colspan="5">&nbsp;Total de cirurgias: <b><%=geral%></b><br><%=agenda_erro%></td>
								<td colspan="5">&nbsp;Total de procedimentos: <b style="color:red;"><%=conta_tipo_2%></b><br><%=agenda_erro%></td>
							<%elseif tipo_relatorio = 3 then%>
								<td colspan="5">&nbsp;Total de itens emprestados: <b style="color:red;"><%=conta_tipo_3%></b><br><%=agenda_erro%></td>
							<%elseif tipo_relatorio = 4 then%>
								<td colspan="5">&nbsp;Total de óticas: <b style="color:red;"><%=conta_tipo_4%></b><br><%=agenda_erro%></td>
							<%end if%>
				</tr>
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" width="100%" height="1"></td></tr>
				<tr class="ok_print"><td colspan="100%"><img src="imagens/px.gif" width="200" height="1"></td></tr>
		<%End if%>
				
		
	<%'end if 'Fim do resultado da pesquisa%>	
		
	<!--tr bgcolor="#C0C0C0">
		<td align="left" colspan="100%" class="no_print">
		
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
		<%=nm_protocolo%>
		</td>
	</tr-->
	<%'end if%>
	<%'end if 'Fim do resultado da pesquisa%>
				<tr><td colspan="100%">&nbsp;</td></tr>
				<tr><td colspan="100%">&nbsp;</td></tr>
			</table></blockquote>
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
</html>


