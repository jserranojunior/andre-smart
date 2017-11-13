<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/numero_extenso.asp"-->
<!--#include file="../includes/util.asp"-->
<!--#include file="../css/geral.htm"-->
<html>
<head>
	<title>::Faturamento::..Adiamento..</title>
<script language="Javascript">
function verifica_data(){
	//hoje = new Date(document.Form.ano_vencimento.value+","+document.Form.mes_vencimento.value+","+document.Form.dia_vencimento.value);
	ori = new Date(document.Form.ori.value);
	ori_dia = ori.getDate()
	ori_mes = ori.getMonth()
	ori_ano = ori.getFullYear()
		
		if (ori_dia < 10 ) ori_dia = '0'+ ori_dia
		if (ori_mes < 10 ) ori_mes = '0'+ ori_mes
			
			data_ori = ori_ano+ori_mes+ori_dia;
	
	novo = new Date(document.Form.ano_vencimento.value,document.Form.mes_vencimento.value,document.Form.dia_vencimento.value);
	var_dia = novo.getDate()
	var_mes = novo.getMonth()
	var_ano = novo.getFullYear()
	
		if (var_dia < 10 ) var_dia = '0'+ var_dia
		if (var_mes < 10 ) var_mes = '0'+ var_mes
			
			data_novo = var_ano+var_mes+var_dia;	
			
		
	//data_ver = document.Form.ano_vencimento.value+document.Form.mes_vencimento.value+document.Form.dia_vencimento.value;
	
	diferenca_datas = data_novo-data_ori
	//alert(diferenca_datas);
	
	if (diferenca_datas < 101){ //&& document.Form.mes_vencimento.value.length > 1) {	
		alert("A data é menor ou igual a data de vencimento original.");
		//document.Form.dif_datas.value="";
		document.Form.mes_vencimento.focus();
		}
	}
	
</script>

	
</head>
<%cd_user = session("cd_codigo")

cd_codigo = request("cd_atrasadas")

if cd_codigo <> "" then
	strsql = "SELECT * FROM TBL_financeiro_fatura_adia_pgto WHERE cd_codigo="&cd_codigo&""
	SET rs_lista = dbconn.execute(strsql)
		if not rs_lista.EOF then
			cd_atrasadas = rs_lista("cd_codigo")
			nm_valor = rs_lista("nm_valor")
			total = int(total+nm_valor)
			cd_unidade = rs_lista("cd_unidade")
			dt_vencimento = rs_lista("dt_vencimento_ori")
			
			dt_vencimento_novo = rs_lista("dt_vencimento_novo")
			dt_dia_venc = day(dt_vencimento_novo)
			dt_mes_venc = month(dt_vencimento_novo)
			dt_ano_venc = year(dt_vencimento_novo)
			
			acao = "adiar_rec_editar"
			data_atual = zero(month(now))&"/"&zero(day(now))&"/"&year(now)
		end if

else
	cd_unidade = request("cd_unidade")
	dt_dia = request("dt_dia")
	dt_mes = request("dt_mes")
	dt_ano = request("dt_ano")
		'dt_vencimento = zero(dt_dia)&"/"&zero(dt_mes)&"/"&dt_ano
		dt_vencimento = zero(dt_dia)&"/"&zero(dt_mes)&"/"&dt_ano
		
	data_atual = zero(month(now))&"/"&zero(day(now))&"/"&year(now)
		dt_dia_venc = dt_dia
		'dt_mes_venc = month(dt_vencimento_novo)
		dt_ano_venc = dt_ano
		
	nm_valor = request("nm_valor")
	acao = "adiar_recebimento"
end if%>
<body>
<table>
	<tr style="color:white; background-color:gray;">
		<td Colspan="2" ><b>Adiar recebimento</b><%'=cd_user%></td>
	</tr>
	<form action="acoes/acoes3.asp" method="post" name="Form" id="Form">
	<input type="hidden" name="cd_user" value="<%=cd_user%>">
	<input type="hidden" name="acao" value="<%=acao%>">
	<input type="hidden" name="dt_vencimento_ori" value="<%=zero(dt_mes)&"/"&zero(dt_dia)&"/"&dt_ano%>">
	<input type="hidden" name="data_atual" value="<%=data_atual%>">
	<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
	<input type="hidden" name="cd_atrasadas" value="<%=cd_atrasadas%>">
	<!--input type="hidden" name="dia_vencimento" value="<%=dt_dia%>">
	<input type="hidden" name="mes_vencimento" value="<%=dt_mes%>">
	<input type="hidden" name="ano_vencimento" value="<%=dt_ano%>"-->
	<input type="hidden" name="nm_valor" value="<%=nm_valor%>">
	<tr><td><input type="hidden" name="dif_datas" value=""></td></tr>
	
	<tr>
		<td>Unidade:<br><img src="../imagens/px.gif" alt="" width="100" height="1" border="0"></td>
		<%if cd_unidade <> "" Then
			xsql = "SELECT * FROM TBL_unidades where cd_codigo = "&cd_unidade&""
			Set rs = dbconn.execute(xsql)
			while not rs.EOF
				strcd_unidade = rs("cd_codigo")
				strnm_unidade = rs("nm_unidade")
				strnm_sigla = rs("nm_sigla")
			rs.movenext
			wend
		end if%>
		<td><b><%=strnm_unidade%></b><br><img src="../imagens/px.gif" alt="" width="150" height="1" border="0"></td>
	</tr>
	<tr>
		<td>Data venc.:</td>
		
		<td><b><%=dt_vencimento%></b><input type="hidden" name="ori" value="<%=month(dt_vencimento)&"/"&day(dt_vencimento)&"/"&year(dt_vencimento)%>"></td>
	</tr>
	<tr>
		<td>Valor:</td>
		<td><b><%=formatnumber(nm_valor,2)%></b></td>
	</tr>
	<tr>
		<td>Novo vencimento:</td>
		<td><input type="text" name="dia_vencimento" size="2" maxlength="2" value="<%=dt_dia_venc%>">/
		<input type="text" name="mes_vencimento" size="2" maxlength="2" value="<%=dt_mes_venc%>" onblur="verifica_data();">/
		<input type="text" name="ano_vencimento" size="4" maxlength="4" value="<%=dt_ano_venc%>">
		</td>
	</tr>
	<tr>
		<td><input type="submit" name="gravar" value="Gravar"></td>
		<td align="right"><img src="../../imagens/ic_del.gif" alt="Apaga funcionario" width="13" height="14" border="0" onclick="javascript:JsDelete('<%=cd_atrasadas%>')"></td></tr>
	</form>
</table>

<SCRIPT LANGUAGE="javascript">
				function JsDelete(cod)
				   {
					if (confirm ("Tem certeza que deseja excluir?"))
				  {
				  document.location.href('acoes/acoes3.asp?cd_atrasadas='+cod+'&acao=adiar_rec_excluir');
					}
					  }
</SCRIPT>
</body>
</html>
