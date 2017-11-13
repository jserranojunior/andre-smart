
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<html>
<head>
	<title>Orçamentos Aprovados</title>
</head>
<!-- #include file="../css/geral.htm"-->
<!--#include file="../js/ajax2.js"-->


<script language="Javascript" TYPE="text/javascript" >

function adiciona_OS(os_int,DEL,val_unit,total_orc,valor,cd_orc,qtd,OS_total){
	if (os_int != '' && val_unit != '' ||DEL != ''){
		//Troca o espaço por cod. Encode
			while (os_int.indexOf(' ') != -1) {
	 		os_int = os_int.replace(' ','%20');}
			while (DEL.indexOf(' ') != -1) {
	 		DEL = DEL.replace(' ','%20');}
			while (val_unit.indexOf(' ') != -1) {
	 		val_unit = val_unit.replace(' ','%20');}
			while (val_unit.indexOf('.') != -1) {
	 		val_unit = val_unit.replace('.',',');}
			
			while (total_orc.indexOf(' ') != -1) {
	 		total_orc = total_orc.replace(' ','%20');}
			while (OS_total.indexOf('$$') != -1) {
	 		OS_total = OS_total.replace('$$','$');}
						
		if (OS_total != ''){
			separador =  '$';
			}
		else
		{
		separador= ''
		}		
	OS_total = OS_total + separador
		if (os_int != ""){
			OS_total = OS_total + os_int + '!' + val_unit + '_'+qtd+'@';
			document.form.OS_result.value=os_int;} //adiciona				
						
		if (DEL != ""){
			OS_total = OS_total.replace(DEL,'');
			document.form.OS_total.value=OS_total;} //subtrai
	
	document.form.OS_total.value=(OS_total.replace("$$", "$"));	
	document.form.OS_total.value=(OS_total); //c
	document.form.OS_interna.value="";
	document.form.valor_unitario.value="";
	document.form.qtd_itens.value="1";
	document.form.OS_interna.focus();
		
		loadXMLDoc("ajax/ajax_OS_lista.asp?OS_total="+OS_total+"&valor_lanc_os="+total_orc+"&valor_total="+valor,function()
			{				
			if (xmlhttp.readyState==4 && xmlhttp.status==200 || xmlhttp.readyState==4 && xmlhttp.status==500)
		    		{
		    			document.getElementById("divOS_lista").innerHTML = xmlhttp.responseText;			
		    		}
		  		}
			);	
	}
	else {
		alert("Falta o nº da O.S. ou valor");
	}
}

function vincula_os(os_valor_qtd){
	//var os_total = document.getElementById("OS_total").value;
	var cd_orcamento = document.getElementById("cd_orcamento").value;
	
	
	loadXMLDoc("acoes/manutencao_acao.asp?filtro=orc_vinc_os&acao=4&cd_orcamento="+cd_orcamento+"&os_total="+os_valor_qtd,function(){
			//alert(xmlhttp.status);		
			if (xmlhttp.readyState==4 && xmlhttp.status==200 || xmlhttp.readyState==4 && xmlhttp.status==500){
		    		document.getElementById("divOS_vinculo").innerHTML = xmlhttp.responseText;			
		    	}
		  	}
		);
	//alert("vinculada-"+cd_orcamento+os_total);
	setInterval(function () { location.reload(forceGet=true); }, 750);
	//location.reload(forceGet=true);
}

function desvincula_os(cd_os){
	if (confirm ("Deseja desvincular a O.S. ao orçamento?")){
	  loadXMLDoc("acoes/manutencao_acao.asp?filtro=orc_vinc_os&acao=5&cd_os="+cd_os,function(){
			//alert(xmlhttp.status);		
			if (xmlhttp.readyState==4 && xmlhttp.status==200 || xmlhttp.readyState==4 && xmlhttp.status==500){
		    		document.getElementById("divOS_lista").innerHTML = xmlhttp.responseText;			
		    	}
		  	}
		);
	//alert("vinculada-"+cd_orcamento+os_total);
	setInterval(function () { location.reload(forceGet=true); }, 250);	
	}
}


</script>
<body>
<%cd_user = session("cd_codigo")
pasta_loc = "manutencao_2\"
arquivo_loc = "orcamentos_aprovados_os.asp"

tipo = request("tipo")
cd_orcamento = request("cd_orcamento")


'***********************************************
'*** Vincula a Ordem de serviço ao Orçamento ***
'***********************************************

%>

<!--#include file="../includes/arquivo_loc.asp"-->
<table align="center" border="1" width="200" style="border-collapse:collapse;">
<tr><td align="center" colspan="10" style="background-color:gray; color:white; font-size:14px;"><b>Orçamento Aprovado</b></td></tr>
<tr style="background-color:silver; font-size:12px;">
	<td colspan="5" align="center">&nbsp;</td>
	<td colspan="3" align="center">Data</td>
<tr style="background-color:silver; font-size:12px;">
	<td>&nbsp;</td>
	<td colspan="3" align="center">Fornecedor<br><img src="../imagens/px.gif" alt="" width="130" height="1" border="0"></td>
	<td align="center">Nº Orç.<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<td align="center">Orç.<br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
	<td align="center">Aprov.<br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
	<td align="center">Entrega<br><img src="../imagens/px.gif" alt="" width="50" height="1" border="0"></td>
	<td align="center">Valor<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
	<td align="center" colspan="2">Parc<br><img src="../imagens/px.gif" alt="" width="70" height="1" border="0"></td>
</tr>

<%num_linha = 1
total_valor = 0

strsql ="SELECT * FROM View_manutencao_orcamento WHERE cd_codigo = '"&cd_orcamento&"' ORDER BY dt_orcamento"
		Set rs_orc = dbconn.execute(strsql)
			if not rs_orc.EOF then
				nm_fornecedor = rs_orc("nm_fornecedor")
				nr_orcamento = rs_orc("nr_orcamento")
				dt_orcamento = rs_orc("dt_orcamento")
					if dt_orcamento <> "" Then	
						dia_orc = zero(day(dt_orcamento))
						mes_orc = zero(month(dt_orcamento))
						ano_orc = right(year(dt_orcamento),2)
						dt_orcamento = dia_orc&"/"&mes_orc&"/"&ano_orc
					end if
					
				dt_aprov_orc = rs_orc("dt_aprov_orc")
					if dt_aprov_orc <> "" Then	
						dia_orc_apv = zero(day(dt_aprov_orc))
						mes_orc_apv = zero(month(dt_aprov_orc))
						ano_orc_apv = right(year(dt_aprov_orc),2)
						dt_aprov_orc = dia_orc_apv&"/"&mes_orc_apv&"/"&ano_orc_apv
					end if
				dt_entrega = rs_orc("dt_entrega")
					if dt_entrega <> "" Then	
						dia_orc_entr = zero(day(dt_entrega))
						mes_orc_entr = zero(month(dt_entrega))
						ano_orc_entr = right(year(dt_entrega),2)
						dt_entrega = dia_orc_entr&"/"&mes_orc_entr&"/"&ano_orc_entr
					end if
				nm_valor = rs_orc("nm_valor")
				nm_valor_orc = nm_valor
					'total_nm_valor = total_nm_valor + nm_valor
					if nm_valor = "" then nm_valor = 0
					
				nr_parcela = rs_orc("nr_parcela")
					if nr_parcela > 1 then
						nr_valor_parcela = nm_valor / nr_parcela
					else
						nr_valor_parcela = 0
					end if
					%>
					<tr>
						<td>&nbsp;</td>
						<td colspan="3"><b><%=nm_fornecedor%></b></td>
						<td align="center"><b><%=nr_orcamento%></b></td>
						<td align="center"><b><%=dt_orcamento%></b></td>
						<td align="center"><b><%=dt_aprov_orc%></b></td>
						<td align="center"><b><%=dt_entrega%></b></td>
						<td align="center">
							<!--a href="javascript:void(0);" onclick="window.open('acoes/manutencao_vinculo_orc_acao.asp?cd_orcamento=<%=cd_orcamento%>&nm_valor_orc=<%=nm_valor_orc%>&dt_orcamento=<%=dt_orcamento%>&dt_aprov_orc=<%=dt_aprov_orc%>&nr_parcela=<%=nr_parcela%>&acao=1', 'orcamento<%=cd_orcamento%>', 'width=587, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"-->
							<b><%if nm_valor <> "" Then response.write(formatNumber(nm_valor,2))%></b><!--/a--></td>
						<%if nr_parcela > 1 then%>
							<td align="center"><b><%=nr_parcela%>X</b></td>
							<td align="right"><b><%if nr_valor_parcela <> "" Then response.write(formatnumber(nr_valor_parcela,2))%></b></td>
						<%else%>
							<td colspan="2">&nbsp;</td>
						<%end if%>
					</tr>
					<tr><td colspan="6">&nbsp;</td></tr>
					<tr style="background-color:silver;">
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td align="center" colspan="2">Item<br><img src="../imagens/px.gif" alt="" width="180" height="1" border="0"></td>
						<td align="center">OS</td>
						<td align="center" colspan="2">Tipo</td>
						<td align="center" colspan="2">Valor</td>
						<td>&nbsp;</td>
					</tr>
					
					<%
					
			'rs_orc.movenext
			end if
		filtro = "orc_cad_upd"
		nm_botao = "Altera"
		

strsql = "SELECT * FROM VIEW_os_lista2 WHERE num_os_assist = '"&nr_orcamento&"' ORDER BY dt_resposta_orc, cd_valor_orc  desc"
'response.write strsql
Set rs = dbconn.execute(strsql)

while not rs.EOF
	cd_os = rs("cd_codigo")
	cd_fornecedor = rs("cd_fornecedor")
	nm_fornecedor = rs("nm_fornecedor")
	cd_equipamento = rs("cd_equipamento")
	nm_equipamento_novo = rs("nm_equipamento_novo")
	cd_unidade = rs("cd_unidade_os")
	num_os = rs("num_os")
	cd_valor_orc = rs("cd_valor_orc")
	if isnumeric(cd_valor_orc) Then valor_total_orcamento = valor_total_orcamento + abs(cd_valor_orc)
	tipo = "orc_aprov"
	cd_gestao = rs("cd_gestao")
	num_os_assist = rs("num_os_assist")
	dt_resposta_orc = rs("dt_resposta_orc")
	dt_retorno = rs("dt_retorno")
		
		if isnumeric(cd_valor_orc) Then total_valor = total_valor + cd_valor_orc
		
			'cd_os%>
	
	<tr>
		<td align="center"><%=num_linha%><%'="-"&dt_retorno%></td>
		<td><%if not isnull(dt_retorno) then response.write("*")%></td>
		<td colspan="2"><%=nm_equipamento_novo%></td>
		<td align="center" colspan="2">
			<%if tipo = "orc_aprov" then%><a href="javascript:void(0);" return false;" onclick="window.open('../manutencao_2/manutencao_andamento2.asp?cd_codigo=<%=cd_os%>&campo=cd_codigo&visual=1&jan=1', 'janela_os<%=cd_os%>', 'width=650, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=zero(cd_unidade)&"."&num_os%></a>
			<%elseif tipo = "financ" then%><a href="javascript:void(0);" return false;" onclick="window.open('../manutencao_2/manutencao_ver_janela2.asp?cd_codigo=<%=cd_os%>&campo=cd_codigo&visual=1&jan=1', 'janela_os<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"><%=zero(cd_unidade)&"."&num_os%></a><%end if%></td>
		<td><%if cd_gestao = 2 then%>Compra<%else%>Manutenção<%end if%></td>
		<td align="right" colspan="2">
			<%if isnumeric(cd_valor_orc) Then%>
				<!--a href="javascript:void(0);" return false;" onclick="window.open('diario_cad3.asp?cd_origem=<%=2&"."&cd_os%>&cd_forn=<%=cd_fornecedor%>&cd_equip=<%=cd_equipamento%>&cd_valor_orc=<%=cd_valor_orc%>&cd_tipo_orc=<%=cd_gestao%>&cd_unidade=<%=cd_unidade%>&campo=cd_codigo&visual=1&jan=1', 'pagamento_<%=cd_os%>', 'width=600, height=430, toolbar=0, location=0, status=0, scrollbars=1,resizable=1'); return false;"-->
					<%=formatnumber(cd_valor_orc,2)%>
				<!--/a-->
			<%else%>
				&nbsp;
			<%end if%></td>
		<td align="center"><img src="../imagens/ic_cancela.gif" alt="Desvincular" width="15" height="15" border="0" onMouseup="desvincula_os(<%=cd_os%>)"></td>
				
	</tr>	
<%
'nm_valor = 0
dt_retorno = ""
ult_num_assist = num_os_assist
num_linha = num_linha + 1
rs.movenext
wend%>
	<tr><td colspan="9" style="font-size:8px;"> &nbsp;( * entregue )</td></tr>
	<tr style="background-color:gray;">
		<td colspan="6">&nbsp;</td>
		<td align="center"><b>Total</b></td>
		<td align="right" style="font-size:11px;" colspan="2"><b><%if valor_total_orcamento <> "" Then response.write(formatnumber(valor_total_orcamento,2))%><%'if valor_total_orcamento <> "" Then response.write(formatnumber(valor_total_orcamento,2))%></b></td>		
	</tr>
</table>

<table align="center" border="1" width="200" style="border-collapse:collapse;">
	<form name="form" action="acoes/manutencao_acao.asp" method="get" id="form">
	<input type="hidden" name="cd_orcamento" id="cd_orcamento" value="<%=cd_orcamento%>" size="40">
	<tr>
		<td colspan="3">
			<p><b>Criar vínculo com O.S. internas</b></p>
		</td>
	<tr>
		<td><b>O.S.</b><br>
			<input type="text" name="OS_interna" size="8" maxlength="10" value="" >
			
			
		</td>
		<td><b>R$ unitario</b><%'=valor_total_orcamento%><br>
			<input type="text" name="valor_unitario" size="8" maxlength="10" value=""> 
		</td>
		<td>
			<b>Qtd itens</b><br>
			<input type="text" name="qtd_itens" size="5" maxlength="10" value="1">
		</td>
		<td>
			<input type="button" name="inclui_OS" value="+" onFocus="adiciona_OS(document.form.OS_interna.value,'',document.form.valor_unitario.value,'<%=valor_total_orcamento%>','<%=nm_valor%>','<%=cd_orcamento%>',document.form.qtd_itens.value,document.form.OS_total.value)">
		</td>
		
	</tr>
	<input type="hidden" name="OS_total" id="OS_total" value="" size="40">	
	<input type="hidden" name="OS_result" value="" size="20">
	<tr>
		<td colspan="4">
			<img src="../imagens/blackdot.gif" width="307" height="1">
		</td>
	<tr>
	<tr>
		<td colspan="4">
			<span id='divOS_lista'><!-- Abre o arquivo: *** ajax/ajax_OS_lista.asp *** -->
				<table>
					<%valor_total_os = valor_total_os + abs(valor_cada)
					
					conta_linha = conta_linha + 1%>
					
					<%diferenca_valor_x_osunica = formatnumber(valor_total_orcamento-nm_valor,2)%>
					<tr><td colspan="4"><img src="../../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td></tr>
					<tr>
						<td colspan="2">Total ordens serviço</td>
						<td align="right"><%if isnumeric(valor_total_orcamento) then response.write(formatnumber(valor_total_orcamento,2))%></td>
					</tr>
					<tr><td colspan="4">&nbsp;</td></tr>
					<tr>
						<td colspan="2">Total orçamento</td>
						<td align="right"><%=nm_valor%></td>
					</tr>
					<tr><td colspan="4"><img src="../../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td></tr>
					<tr>
						<td colspan="2">Diferença</td>
						<td align="right" style="color:red;"><b><%=diferenca_valor_x_osunica%></b></td>
					</tr>
				</table>
			</span>
		</td>
	</tr>
	</form>
			
</table>
</body>
</html>
