<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%
strcd_unidade = request("cd_unidade")
str_status = request("status")
acao = request("acao")
	if acao = "" Then
	acao = "inserir"
	Else
	acao = "editar"
	end if
%>

<head>
	<title>Cadastro de nomes de itens</title>
</head>
<script language="JavaScript">
	function validar_unidade(shipinfo){
			if (shipinfo.nm_unidade.value.length==""){window.alert ("O nome da unidade deve ser digitado.");shipinfo.nm_unidade.focus();return (false);}
			if (shipinfo.nm_sigla.value.length==""){window.alert ("A sigla da unidade deve ser digitada.");shipinfo.nm_sigla.focus();return (false);}
			if (shipinfo.cd_unidade_tipo.value.length==""){window.alert ("O tipo da unidade deve ser informado.");shipinfo.cd_unidade_tipo.focus();return (false);}
			if (shipinfo.cd_empresa.value.length==""){window.alert ("Informe a empresa prestadora de serviços.");shipinfo.cd_empresa.focus();return (false);}
			if (shipinfo.dt_i_contrato_dia.value.length==""){window.alert ("Informe a data do início do contrato.");shipinfo.dt_i_contrato_dia.focus();return (false);}
			if (shipinfo.dt_i_contrato_mes.value.length==""){window.alert ("Informe a data do início do contrato.");shipinfo.dt_i_contrato_mes.focus();return (false);}
			if (shipinfo.dt_i_contrato_ano.value.length==""){window.alert ("Informe a data do início do contrato.");shipinfo.dt_i_contrato_ano.focus();return (false);}
			
			//if (shipinfo.existe.value.length==""){window.alert ("Digite um nome diferente.");shipinfo.nm_equipamento.focus();return (false);}
		return (true);}	
</script>

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

<!--#include file="../js/ajax.js"-->
<!--#include file="../ferramentas/js/ferramentas.js"-->

<body onload="foco()">
<table class="txt" align="center" id="no_print" width="300" style="font-family:verdana; font-size:10px;">
<%if strcd_unidade <> "" Then
	xsql = "up_unidade_busca @cd_unidade='"&strcd_unidade&"', @nm_unidade=' ', @cd_status='"&str_status&"'"
	set rs_unid = dbconn.execute(xsql)
	
		if not rs_unid.EOF then
			cd_unidade = rs_unid("cd_codigo")
			nm_unidade = rs_unid("nm_unidade")
			nm_unidade_nome = rs_unid("nm_unidade_nome")
			nm_sigla = rs_unid("nm_sigla")
			nm_endereco = rs_unid("nm_endereco")
				if nm_endereco <> "" then
					sep1 = instr(1,nm_endereco,"!",1)
					sep2 = instr(1,nm_endereco,"#",1)
					sep3 = instr(1,nm_endereco,"$",1)
					sep4 = instr(1,nm_endereco,"[",1)
					sep5 = instr(1,nm_endereco,"]",1)
					
					
					nm_rua = mid(nm_endereco,1,sep1-1)
					nm_no = mid(nm_endereco,sep1+1,sep2-sep1-1)
					nm_cep = mid(nm_endereco,sep2+1,((sep3-sep2))-1)
						sepcep = instr(1,nm_cep,"-",1)
						if nm_cep <> "" Then
							nm_cep1 = mid(nm_cep,1,sepcep-1)
							nm_cep2 = mid(nm_cep,sepcep+1,len(nm_cep)-sepcep)
						end if
					nm_bairro = mid(nm_endereco,sep3+1,((sep4-sep3))-1)
					nm_cidade = mid(nm_endereco,sep4+1,((sep5-sep4))-1)
					nm_estado = mid(nm_endereco,sep5+1,((len(nm_endereco)-sep5)))
				end if
			nm_contato = rs_unid("nm_contato")
			nm_telefone = rs_unid("nm_telefone")
			nm_cnpj = rs_unid("nm_cnpj")
			nm_ie = rs_unid("nm_ie")
			cd_hospital = rs_unid("cd_hospital")
			cd_cliente_empresa = rs_unid("cd_cliente_empresa")
			dt_inicio_contrato = rs_unid("dt_inicio_contrato")
				dt_i_contrato_dia = zero(day(dt_inicio_contrato))
				dt_i_contrato_mes = zero(month(dt_inicio_contrato))
				dt_i_contrato_ano = year(dt_inicio_contrato)
			dt_fim_contrato = rs_unid("dt_fim_contrato")
				dt_f_contrato_dia = zero(day(dt_fim_contrato))
				dt_f_contrato_mes = zero(month(dt_fim_contrato))
				dt_f_contrato_ano = year(dt_fim_contrato)
			nm_cor = trim(rs_unid("nm_cor"))
			cd_status = rs_unid("cd_status")
			cd_prazo_faturamento = rs_unid("cd_prazo_faturamento")
		end if
end if%>
	
	<tr bgcolor="#ffcc00">
		<td align="center" colspan="100%" style="color:red;font-size=13px;"><b>Cadastro de Unidades</b></td>
	</tr>
	<form action="acao/cadastros_acao.asp" method="post" name="Form" onsubmit="return validar_unidade(document.Form);">
	<!--input type="hidden" name="janela" value="1"-->
	<input type="hidden" name="modo" value="unidade">
	<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
	<input type="hidden" name="acao" value="<%=acao%>">
	<!--input type="hidden" name="status" value="<%'=cd_status%>"-->
	<input type="hidden" name="jan" value="1">
	<%if str_status = 0 then
		atividade = "#ccffff"
	Elseif str_status = 1 then
		atividade = "#ffcc99"
	end if%>	
	<tr bgcolor=#cococo>
		<td align=center colspan="4"><img src="imagens/px.gif"  height=1></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Nome Real:</td>
	    <td colspan="3"><input type="text" name="nm_unidade_nome" size="60" value="<%=nm_unidade_nome%>" class="inputs"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Unidade:</td>
	    <td colspan="3"><input type="text" name="nm_unidade" size="60" value="<%=nm_unidade%>" class="inputs"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Sigla:</td>
	    <td colspan="3"><input type="text" name="nm_sigla" size="15" value="<%=nm_sigla%>" class="inputs" maxlength="5"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Endereço:</td>
	    <td colspan="3"><input type="text" name="nm_endereco" size="51" value="<%=nm_rua%>" class="inputs">,<input type="text" name="nm_numero" size="3" value="<%=nm_no%>" class="inputs"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; CEP:</td>
	    <td><input type="text" name="nm_cep1" size="6" value="<%=nm_cep1%>" class="inputs" maxlength="5" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 5)" onFocus="PararTAB(this);"> <input type="text" name="nm_cep2" size="3" value="<%=nm_cep2%>" class="inputs" maxlength="3"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 3)" onFocus="PararTAB(this);"></td>
		<td>&nbsp; Bairro:</td>
	    <td><input type="text" name="nm_bairro" size="20" value="<%=nm_bairro%>" class="inputs"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Cidade:</td>
	    <td><input type="text" name="nm_cidade" size="20" value="<%=nm_cidade%>" class="inputs"></td>
		<td>&nbsp; UF:</td>
		<td><input type="text" name="nm_estado" size="3" value="<%=nm_estado%>" class="inputs" maxlength="2" onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Contato:</td>
	    <td><input type="text" name="nm_contato" size="20" value="<%=nm_contato%>" class="inputs"></td>
		<td>&nbsp; Telefone:</td>
		<td><input type="text" name="nm_telefone" size="20" value="<%=nm_telefone%>" class="inputs" maxlength="50"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; CNPJ:</td>
	    <td><input type="text" name="nm_cnpj" size="20" value="<%=nm_cnpj%>" class="inputs"></td>
		<td>&nbsp; I.E.:</td>
		<td><input type="text" name="nm_ie" size="20" value="<%=nm_ie%>" class="inputs"></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Tipo:</td>
	    <td>
			<select name="cd_unidade_tipo">
						<option value=""></option>
						<%xsql = "up_unidade_tipo_lista"
						set rs = dbconn.execute(xsql)
							while not rs.EOF
								cd_unidade_tipo = rs("cd_unidade_tipo")
								nm_unidade_tipo = rs("nm_unidade_tipo")%>
								<option value="<%=cd_unidade_tipo%>" <%if cd_hospital = cd_unidade_tipo then%>SELECTED<%end if%>><%=nm_unidade_tipo%></option>
							<%rs.movenext
							wend%>
			</select>
		</td>
		<td>&nbsp; Empresa:</td>
		<td><select name="cd_empresa">
						<option value="" SELECTED></option>
						<%xsql = "up_empresa_lista"
						set rs = dbconn.execute(xsql)
							while not rs.EOF
							cd_empresa = rs("cd_codigo")%>
								<option value="<%=cd_empresa%>" <%if cd_cliente_empresa = cd_empresa then%>SELECTED<%end if%>><%=rs("nm_empresa")%></option>
							<%rs.movenext
							wend%>
			</select></td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Contrato:</td>
	    <td><input type="text" name="dt_i_contrato_dia" size="2" maxlength="2" value="<%=dt_i_contrato_dia%>" class="inputs"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_i_contrato_mes" size="2" maxlength="2" value="<%=dt_i_contrato_mes%>" class="inputs"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_i_contrato_ano" size="4" maxlength="4" value="<%=dt_i_contrato_ano%>" class="inputs"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
		<%if cd_unidade <> "" Then%>
			<td>&nbsp; Recisão:</td>
			<td><input type="text" name="dt_f_contrato_dia" size="2" maxlength="2" value="<%=dt_f_contrato_dia%>" class="inputs"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_f_contrato_mes" size="2" maxlength="2" value="<%=dt_f_contrato_mes%>" class="inputs"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);">/<input type="text" name="dt_f_contrato_ano" size="4" maxlength="4" value="<%=dt_f_contrato_ano%>" class="inputs"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);"></td>
		<%else%>
			<td colspan="2">&nbsp;</td>
		<%end if%>
	</tr>
	<tr bgcolor="<%=atividade%>">
		<td>&nbsp; Cor:</td>
		<td><select name="nm_cor">
				<option value=""></option>
				<option value="ffff99" <%if nm_cor = "ffff99" then%> SELECTED <%end if%> style="background-color:#ffff99;">Amarelo claro</option>
				<option value="ffff00" <%if nm_cor = "ffff00" then%> SELECTED <%end if%> style="background-color:#ffff00;">Amarelo</option>
				<option value="ffcc00" <%if nm_cor = "ffcc00" then%> SELECTED <%end if%> style="background-color:#ffcc00;">Amarelo escuro</option>
				<option value="ffa264" <%if nm_cor = "ffa264" then%> SELECTED <%end if%> style="background-color:#ffa264;">Laranja claro</option>
				<option value="ff9900" <%if nm_cor = "ff9900" then%> SELECTED <%end if%> style="background-color:#ff9900;">Laranja</option>
				<option value="ff6600" <%if nm_cor = "ff6600" then%> SELECTED <%end if%> style="background-color:#ff6600;">Laranja escuro</option>
				<option value="ff80ff" <%if nm_cor = "ff80ff" then%> SELECTED <%end if%> style="background-color:#ff80ff;">Rosa claro</option>
				<option value="ff00cc" <%if nm_cor = "ff00cc" then%> SELECTED <%end if%> style="background-color:#ff00cc;">Rosa</option>
				<option value="dd33aa" <%if nm_cor = "dd33aa" then%> SELECTED <%end if%> style="background-color:#dd33aa;">Rosa escuro</option>
				<option value="ff6666" <%if nm_cor = "ff6666" then%> SELECTED <%end if%> style="background-color:#ff6666;">Vermelho claro</option>
				<option value="ff0000" <%if nm_cor = "ff0000" then%> SELECTED <%end if%> style="background-color:#ff0000;">Vermelho</option>
				<option value="990033" <%if nm_cor = "990033" then%> SELECTED <%end if%> style="background-color:#990033;">Vermelho escuro</option>
				<option value="bb9ed6" <%if nm_cor = "bb9ed6" then%> SELECTED <%end if%> style="background-color:#bb9ed6;">Roxo claro</option>
				<option value="9900ff" <%if nm_cor = "9900ff" then%> SELECTED <%end if%> style="background-color:#9900ff;">Roxo</option>
				<option value="660099" <%if nm_cor = "660099" then%> SELECTED <%end if%> style="background-color:#660099;">Roxo escuro</option>
				<option value="a8ddfd" <%if nm_cor = "a8ddfd" then%> SELECTED <%end if%> style="background-color:#a8ddfd;">Azul claro</option>
				<option value="0000ff" <%if nm_cor = "0000ff" then%> SELECTED <%end if%> style="background-color:#0000ff;">Azul</option>
				<option value="00366c" <%if nm_cor = "00366c" then%> SELECTED <%end if%> style="background-color:#00366c;">Azul escuro</option>
				<option value="66ff00" <%if nm_cor = "66ff00" then%> SELECTED <%end if%> style="background-color:#66ff00;">Verde claro</option>
				<option value="4f9d00" <%if nm_cor = "4f9d00" then%> SELECTED <%end if%> style="background-color:#4f9d00;">Verde</option>
				<option value="006633" <%if nm_cor = "006633" then%> SELECTED <%end if%> style="background-color:#006633;">Verde escuro</option>
				<option value="996600" <%if nm_cor = "996600" then%> SELECTED <%end if%> style="background-color:#996600;">Marrom</option>
				<option value="bbbbbb" <%if nm_cor = "bbbbbb" then%> SELECTED <%end if%> style="background-color:#bbbbbb;">Cinza claro</option>
				<option value="747474" <%if nm_cor = "747474" then%> SELECTED <%end if%> style="background-color:#747474;">Cinza médio</option>
				<option value="3c3c3c" <%if nm_cor = "3c3c3c" then%> SELECTED <%end if%> style="background-color:#3c3c3c;">Cinza escuro</option>
			</select>
		</td>
		<%if cd_unidade <> "" Then%>
			<td>&nbsp;Status:</td>
			<td><select name="cd_status">
					<option value="0" <%if cd_status = "0" then%>SELECTED<%end if%>>Inativo</option>
					<option value="1" <%if cd_status = "1" then%>SELECTED<%end if%>>Ativo</option>
				</select>
			</td>
		<%else%>
			<td colspan="2"></td>
		<%end if%>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp; Prazo<br> &nbsp;&nbsp;Faturamento:</td>
	    <td><input type="text" name="cd_prazo_faturamento" size="2" maxlength="2" value="<%=cd_prazo_faturamento%>" class="inputs"> (qtd meses)</td>
		<td>&nbsp;</td>
	    <td>&nbsp;</td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td colspan="4">&nbsp;</td>
	</tr>
	<tr bgcolor="<%=atividade%>">
	    <td>&nbsp;</td>
		<td align="left" colspan="3"> &nbsp; <span id='divEquip_novo'><%'if acao = "editar" then%><input type="submit" name="envia" value="<%if acao = "editar" then%>Altera<%else%>Cadastra<%end if%>"><%'end if%></span></td>
	</tr>
	<tr>
		<td><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="160" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="70" height="1" border="0"></td>
		<td><img src="../imagens/blackdot.gif" alt="" width="160" height="1" border="0"></td>
	</tr>
</table>
<!--/td></tr></table-->
<SCRIPT LANGUAGE="javascript">
function JsDelete(cod1)
{
  if (confirm ("Tem certeza que deseja excluir?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=2&acao=excluir');
	  }
}
function JsDesativar(cod1)
{
  if (confirm ("Tem certeza que deseja tornar o equipamento inativo?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=0&acao=excluir');
	  }
}
function JsReativar(cod1)
{
  if (confirm ("Tem certeza que deseja reativar o cadastro do equipamento?"))
	  {
		document.location.href('acoes/cadastros_acao.asp?tipo=equipamento&cd_equipamento='+cod1+'&modo=equipamento&status=0&acao=editar');
	  }
}
</SCRIPT>
</body>
</html>
