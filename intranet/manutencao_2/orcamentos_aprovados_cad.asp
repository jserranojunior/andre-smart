
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../css/geral.htm"-->
<!--include file="../manutencao_2/js/manutencao2.js"-->

<!--#include file="../js/ajax2.js"-->


<script language="Javascript" TYPE="text/javascript" >

function adiciona_OS(a,b,c,d,OS_total){
	a=a; 
	b=b;
	c=c;
	d=d;
	OS_total=OS_total;
		
		//Troca o espaço por cod. Encode
			while (a.indexOf(' ') != -1) {
	 		a = a.replace(' ','%20');}
			while (b.indexOf(' ') != -1) {
	 		b = b.replace(' ','%20');}
			while (c.indexOf(' ') != -1) {
	 		c = c.replace(' ','%20');}
			while (d.indexOf(' ') != -1) {
	 		d = d.replace(' ','%20');}
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
		if (a != ""){
			OS_total = OS_total + a + '!' + c + '_'+ d + '@';
			document.form.OS_result.value=a;} //adiciona				
						
		if (b != ""){
			OS_total = OS_total.replace(b,'');
			document.form.OS_total.value=OS_total;} //subtrai
	
	document.form.OS_total.value=(OS_total.replace("$$", "$"));	
	document.form.OS_total.value=(OS_total); //c
	document.form.OS_interna.value="";
	document.form.valor_unitario.value="";
	document.form.OS_interna.focus();
	
	
		//loadXMLDoc("ajax/ajax_OS_lista.asp?OS_total=xxx+"+OS_total,function()
		  						
		loadXMLDoc("ajax/ajax_OS_lista.asp?OS_total="+OS_total,function()
			{				
			//alert(xmlhttp.readyState+'-'+xmlhttp.status);
			if (xmlhttp.readyState==4 && xmlhttp.status==200)
		    		{
		    			document.getElementById("divOS_lista").innerHTML = xmlhttp.responseText;						
		    		}
		  		}
			);	
}
</script>

<%cd_user = session(cd_codigo)
pasta_loc = "manutencao2\"
arquivo_loc = "orcamentos_aprovados_cad.asp"

tipo = request("tipo")
cd_orcamento = request("cd_orcamento")%>
<html>
<head>
	<title>Orç. Aprovado</title>
</head>

<body>
<table align="center" style="background-color:#c6e2c2;">
<!--#include file="../includes/arquivo_loc.asp"-->
		<tr>
			<td align="center" colspan="4" style="background-color:gray; color:white; font-">Cadastrar Orçamentos Aprovados</td>
		</tr>
		<%if cd_orcamento <> "" Then
			strsql ="SELECT * FROM View_manutencao_orcamento WHERE cd_codigo = '"&cd_orcamento&"' ORDER BY dt_orcamento"
			Set rs_orc = dbconn.execute(strsql)
				while not rs_orc.EOF
					nm_fornecedor = rs_orc("nm_fornecedor")
					nr_orcamento = rs_orc("nr_orcamento")
					dt_orcamento = rs_orc("dt_orcamento")
						if dt_orcamento <> "" Then	
							dia_orc = zero(day(dt_orcamento))
							mes_orc = zero(month(dt_orcamento))
							ano_orc = year(dt_orcamento)
						end if
					
					dt_aprov_orc = rs_orc("dt_aprov_orc")
						if dt_aprov_orc <> "" Then
							dia_aprov_orc = zero(day(dt_aprov_orc))
							mes_aprov_orc = zero(month(dt_aprov_orc))
							ano_aprov_orc = year(dt_aprov_orc)
						end if
					
					nm_valor = rs_orc("nm_valor")
						'total_nm_valor = total_nm_valor + nm_valor
					nr_parcela = rs_orc("nr_parcela")
						if nr_parcela > 1 then
							nr_valor_parcela = nm_valor / nr_parcela
						else
							nr_valor_parcela = 0
						end if
				rs_orc.movenext
				wend
			filtro = "orc_cad_upd"
			nm_botao = "Altera"
		else
			filtro = "orc_cad"
			nm_botao = "Cadastra"
		end if%>
		<tr>
			<td>&nbsp;
			<form name="form" action="acoes/manutencao_acao.asp" method="get" id="form">
			<input type="hidden" name="filtro" value="<%=filtro%>">
			<input type="hidden" name="cd_orcamento" value="<%=cd_orcamento%>">
			<input type="hidden" name="acao" value="1">
			<input type="hidden" name="jan" value="0">
			</td>
		</tr>
		<tr>
			<td><b>Nº Orç.</b><br>
				<img src="../imagens/px.gif" width="120" height="1"><br>
				<input type="text" name="nr_orcamento" size="10" maxlength="20" value="<%=nr_orcamento%>" onKeyup="ver_orcamento();"></td>
			<%strsql ="TBL_Fornecedor order by nm_fornecedor"
			Set rs_forn = dbconn.execute(strsql)%>
			<td colspan="2"><b>Assist./Fornecedor.</b><br>
			<img src="../imagens/px.gif" width="120" height="1"><br>
				<span id='divForn'>
					<select name="cd_fornecedor" class="inputs" onFocus="nextfield ='observacoes';">
						<option value="">
						<%Do while Not rs_forn.EOF%><%fornecedor=rs_forn("nm_fornecedor")%><%if nm_fornecedor=fornecedor Then%><%ck_forn="selected"%><%else ck_forn=""%><%end if%>
						<option value="<%=rs_forn("cd_codigo")%>" <%=ck_forn%>><%=rs_forn("nm_fornecedor")%></option><%rs_forn.movenext
						Loop%>
					</select>
				</span>
			</td>
							
		</tr>
		<tr>
			<td><b>Data</b><br>
				<img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_dia" size="2" maxlength="2" value="<%=dia_orc%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
				<input type="text" name="dt_mes" size="2" maxlength="2" value="<%=mes_orc%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
				<input type="text" name="dt_ano" size="4" maxlength="4" value="<%=ano_orc%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
			<td><b>valor</b><br>
				<img src="../imagens/px.gif" width="120" height="1"><br>
				<input type="text" name="nm_valor" size="10" maxlength="20" value="<%=nm_valor%>" onKeyup="valores();" style="text-align:right"></td>
			<td><b>Qtd. Parcelas.</b><br>
				<img src="../imagens/px.gif" width="80" height="1"><br>
				<input type="text" name="nr_parcela" size="2" maxlength="3" value="<%=nr_parcela%>">
				
				<%=cd_user%></td>
		</tr>
		<tr>
			<td><b>Data aprovação</b><br>
				<img src="../imagens/px.gif" width="170" height="1"><br>
				<input type="text" name="dt_dia_aprov" size="2" maxlength="2" value="<%=dia_aprov_orc%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
				<input type="text" name="dt_mes_aprov" size="2" maxlength="2" value="<%=mes_aprov_orc%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/
				<input type="text" name="dt_ano_aprov" size="4" maxlength="4" value="<%=ano_aprov_orc%>" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);">
			</td>
		</tr>
		<!--tr>
			<td colspan="3">
				<img src="../imagens/blackdot.gif" width="460" height="1"><br>
				<p><b>Criar vínculo com O.S. internas</b></p>
			</td>
		<tr>
			<td><b>O.S.</b><br>
				<input type="text" name="OS_interna" size="8" maxlength="10" value="">
				
				
			</td>
			<td><b>R$ unitario</b><br>
				<input type="text" name="valor_unitario" size="8" maxlength="10" value=""> 
				<td><input type="button" name="inclui_OS" value="+" onFocus="adiciona_OS(document.form.OS_interna.value,'',document.form.valor_unitario.value,'1',document.form.OS_total.value)">
			</td>
			
		</tr-->
		<input type="hidden" name="OS_total" value="" size="40">	
		<input type="hidden" name="OS_result" value="" size="20">
		<tr>
			<td colspan="3">
				<img src="../imagens/blackdot.gif" width="460" height="1">
			</td>
		<tr>
		<tr>
			<td><span id="ver_orcamento"><input type="button" value="Gravar" onclick="javascript:submit();"></span></td>
			<td><input type="button" name="Reset" value="Reset" onclick="javascript:window.location.replace('orcamentos_aprovados_cad.asp?tipo=cad_orc_aprov')"></td>
		</tr>
		<tr>
			<td colspan="3"><span id='divOS_lista'></span>
			</td>
		</tr>
			</form>
	
	</table>


</body>
</html>
<SCRIPT LANGUAGE="javascript">
		formataMoeda(document.forms.form.nm_valor, 2);
</SCRIPT>