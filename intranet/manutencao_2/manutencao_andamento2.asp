<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/util.asp"--> 
<!--#include file="../includes/estilos_css.htm"-->
<!--#include file="../js/ajax.js"-->
<!--#include file="../js/ajax2.js"-->
<!--#include file="../manutencao_2/js/manutencao.js"-->
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<%esconder = "1"%>
<BODY>
<br> 

<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "dt_manut_orc"; // nome do primeiro campo do site
netscape = "";
ver = navigator.appVersion; len = ver.length;
for(iln = 0; iln < len; iln++) if (ver.charAt(iln) == "(") break;
	netscape = (ver.charAt(iln+1).toUpperCase() != "C");
	function keyDown(DnEvents) {
		// ve quando e o netscape ou IE
		k = (netscape) ? DnEvents.which : window.event.keyCode;
		if (k == 13) { // preciona tecla enter
		if (nextfield == 'done') {
			//alert("viu como funciona?");
			eval('document.form.' + nextfield + '.focus()');
			return false;
			//return true; // envia quando termina os campos
		} else {
			// se existem mais campos vai para o proximo
			eval('document.form.' + nextfield + '.focus()');
			return false;
		}
	}
}
document.onkeydown = keyDown; // work together to analyze keystrokes
if (netscape) document.captureEvents(Event.KEYDOWN|Event.KEYUP);

function valores(){
	var qtd = document.form.cd_qtd_valor.value;
	var valor = document.form.cd_valor_unit.value;
	
	var valor = valor;
	var	valor = valor.replace('.','');
	var	valor = valor.replace(',','.');
	//var qtd = parseFloat(qtd);
	
	if (qtd == ''){
		//var valor_total = parseFloat(valor*1);
		var valor_total = valor;
		//document.form.cd_qtd_valor.value = 1;
		}
	else{
		//var valor_total = parseFloat(valor*qtd);
		
		/*var	valor = valor.replace('.','');
		var	valor = valor.replace(',','.');*/
		var valor_total = valor*qtd;
		//var valor_total = valor;
		}
		
		//document.form.cd_valor_orc.value = valor_total.toFixed(2);
		document.form.cd_valor_orc.value = valor_total;
		}
		//return !Number(valor_total) ? valor_total : Number(valor_total)%1 === 0 ? Number(valor_total).toFixed(2) : valor_total;

		
function mostra_orcamento()
 {
 	var num = form.num_os_assist.value;
	//alert(num);
     var oHTTPRequest_orc = createXMLHTTP(); 
    oHTTPRequest_orc.open("post", "../manutencao_2/ajax/ajax_orcamento_mostra.asp", true);
     oHTTPRequest_orc.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
     oHTTPRequest_orc.onreadystatechange=function(){
      if (oHTTPRequest_orc.readyState==4){
         document.all.divorcamento.innerHTML = oHTTPRequest_orc.responseText;}}
       oHTTPRequest_orc.send("num_os_assist=" + num);
 }
 
 function mostra_orcamento2(num_orc,cd_forn)
 {
 	loadXMLDoc("../manutencao_2/ajax/ajax_orcamento_mostra2.asp?num_os_assist="+num_orc+"&cd_fornecedor="+cd_forn,function()
			{
			if (xmlhttp.readyState==4 && xmlhttp.status==200 || xmlhttp.readyState==4 && xmlhttp.status==500)
		    		{
		    			document.getElementById("divorcamento").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
			
		// *** preenche o codigo do or�amento ***
		
}		
	
function alimenta_info_orcamento(cd_codigo,dt_resposta_orc,dt_aprov_orc){
			document.getElementById("codOrc").value=cd_codigo;
			document.getElementById("dtOrca").value=dt_resposta_orc;
			document.getElementById("dtAprov").value=dt_aprov_orc;
 }
 
 function mostra_pat(tipo,codigo_pat){
		//*** Seleciona o numero do patrimonio ou o numero de serie ***
		loadXMLDoc("manutencao_2/ajax2/ajax_os_pat_NS_mostra.asp?tipo_pat="+tipo+"&no_patrimonio="+codigo_pat+"&no_serie=",function()
			{
			if (xmlhttp.readyState==4 && xmlhttp.status==200 || xmlhttp.readyState==4 && xmlhttp.status==500)
		    		{
		    			document.getElementById("divPatNS_mostra").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
		// limpa campos do equipamento/ otica/ instrumento
			document.getElementById("numNS").value='';
			document.getElementById("cd_patrimonio").value='';
			document.getElementById("cd_equipamento").value='';
			document.getElementById("cd_marca").value='';
			document.getElementById("cd_especialidade").value='';
			document.getElementById("cd_ns").value='';
			document.getElementById("num_qtd").value='';
	}
  
  // End -->
</script>

<%
cd_user = session("cd_codigo")
pasta_loc = "Manutencao2\"
arquivo_loc = "manutencao_andamento2.asp"

usuario_restrito = int("650")
jan = request("jan")

cd_codigo = request("cd_codigo")
partic  = request("partic")
cd_und = request("cd_prefixo")
num_os = request("num_os")
	
	if partic = "busca" then
		cd_codigo = cd_codigo & " AND cd_unidade= "&cd_und
	end if
	
campo = request("campo")
ordem = request("ordem")
list = request("list")
lista = "1"
strvisual = request("visual")

if strvisual = "1" Then
'visual = "hidden"
visual = "visible"
Else
visual = "visible"
End if

if ordem = "" Then
ordem = "1"
Elseif ordem = "1" Then
ordem = "0"
End if

	if ordem = "0" Or ordem = "" Then
	ordem = "DESC"
	Else
	ordem = "ASC"
	End if

	If campo= "" AND cd_codigo = "" Then
	campo = "num_os"
	formulario = "0"
	End If

	compare = "="
	strtop = "100 percent"
	outros = ""
%>
<!--#include file="../includes/arquivo_loc.asp"-->
<%
xsql ="up_os_lista2 @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
Set rs = dbconn.execute(xsql)
If rs.EOF Then%>


			<table width="550" border="1" cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
				<tr>		
					<td class="txt_cinza" colspan="6">
					<b>Manuten��o &raquo; - <font color="red">Listagem1</font></b><br><br></td>					
					</td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao2.asp" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova">
					<tr id="no_print"><td colspan="100%">NOVA O.S.:  <input type="text" name="cd_unidade" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"><input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%'if cd_user <> usuario_restrito then
				if jan <> 1 then%>
				<tr style="visibility:<%=visual%>;">
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao2.asp?tipo=pendencias&filtro=pendentes">Pendentes</a> 
						/ <a href="manutencao2.asp?tipo=pendencias&filtro=encerradas&strtop=20">Encerradas</a>
						/ <a href="manutencao2.asp?tipo=pendencias&filtro=todas&strtop=20">Todas2</a>
						/ <a href="manutencao2.asp?tipo=relatorio">Buscas</a>
					</td>
				</tr>
				<tr><td><br>*** N�o foi poss�vel encontrar a O.S. solicitada ***<br>&nbsp;</td></tr>
				<%end if
			registros = 0
			Else
			Do while not rs.EOF
				cd_cod_andamento = rs("cd_codigo")
				cd_unidade = rs("cd_unidade")
				num_os = rs("num_os")
				cd_natureza_os = rs("cd_natureza_os")
					if cd_natureza_os = "" then
						 cd_natureza_os = "0"
					end if
					
				dt_os = rs("dt_os")
				'nm_unidade = rs("nm_unidade")
				nm_unidade = rs("nm_sigla")
				nm_especialidade = rs("nm_especialidade")
				nm_equipamento = rs("nm_equipamento_novo")
			
				dt_entrada = rs("dt_entrada")
				cd_avaliacao = rs("cd_avaliacao")
				nm_fornecedor = rs("nm_fornecedor")
				cd_fornecedor = rs("cd_fornecedor")
				cd_liberacao = rs("cd_liberacao")
				
				cd_user_cad = rs("cd_user_cad")
				dt_cad = rs("dt_cad")
				cd_user_upd = ("cd_user_upd")
				dt_upd = ("dt_upd")
				
				
				
				
				
				if cd_user_cad <> "" Then
				
				dt_cad = zero(day(dt_cad))&"/"&zero(month(dt_cad))&"/"&right(year(dt_cad),2)
				dt_os = zero(day(dt_os))&"/"&zero(month(dt_os))&"/"&right(year(dt_os),2)
				dt_entrada = zero(day(dt_entrada))&"/"&zero(month(dt_entrada))&"/"&right(year(dt_entrada),2)
				
				
					
					<!-- Formata��o das datas -->
				end if
				
				if cd_user_upd <> "" Then
					
				end if
				
				cd_codigo = rs("cd_codigo")
			
				num_qtd = rs("num_qtd")
				nm_marca = rs("nm_marca")
				cd_ns = rs("cd_ns")
				cd_patrimonio = rs("cd_patrimonio")
				no_patrimonio = rs("no_patrimonio")
				no_pat_old = rs("no_pat_old")
				motivo = rs("motivo")
				nm_solicitante = rs("nm_solicitante")
				observacoes = rs("observacoes")
				sequencia = "1"
				fecha = rs("fecha")
					if int(fecha) = 1 then
						bg_cor = "ffffcc"
					else
						bg_cor = "f5f5f5"
					end if
					
				
					'****** Mostra os nomes das avaliacaoes ***
						strsql ="SELECT * FROM TBL_avaliacao WHERE cd_codigo='"&cd_avaliacao&"'"' up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
						Set rs_aval = dbconn.execute(strsql)
						Do while not rs_aval.EOF
						nm_avaliacao = rs_aval("nm_avaliacao")
						rs_aval.movenext
						loop
					'********************************************
				rs.movenext
				loop
			
			
			css_tabela = "border:2px solid gray; border-collapse:collapse; width:670px; margin-left: auto; margin-right:auto;"%>
			<table border="1" style="<%=css_tabela%>">
				<tr>		
					<td class="txt_cinza" colspan="6">
						<b>Manuten��o &raquo; - <font color="red">Listagem</font></b><b style="color=white;"><%=cd_cod_andamento%> <%'=usuario_restrito%></b><br><br></td>
					</td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao2.asp?tipo=nova" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova">
					<tr id="ok_print"><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%'if cd_user <> usuario_restrito then
				if jan <> 1 then%>
				<tr style="visibility:<%=visual%>;">
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao2.asp?tipo=pendencias&filtro=pendentes">Pendentes</a> 
						/ <a href="manutencao2.asp?tipo=pendencias&filtro=encerradas&strtop=20">Encerradas</a>
						/ <a href="manutencao2.asp?tipo=pendencias&filtro=todas&strtop=20">Todas</a>
						/ <a href="manutencao2.asp?tipo=relatorio">Buscas</a> /
					</td>
				</tr>
				<%end if%>				
			<!--In�cio da etapa 1.1-->
			<%jan=request("jan")
			if jan = 1 then
				caminho = "../"
			else
				jan = 0
			end if%>
				<tr>			
					<form method="post" action="<%=caminho%>manutencao_2/acoes/manutencao_acao.asp" name="form" id="form" onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="filtro" value="andamento">
						<input type="hidden" name="jan" value="<%=jan%>">
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
				<tr bgcolor="#b4b4b4">
					<td  colspan="100%" align="center" class="textopadrao">SOLICITA��O </td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>O.S.</b></td>
					<td><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></td>
					<!--td align="center"><b>Natureza</b></td-->
					<td align="center"><b><%if cd_natureza_os = 1 then%>Manuten��o<%elseif cd_natureza_os = 2 then%>Compra<%end if%></b></td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Data</b></td>
					<td><%=dt_os%></td>
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Solicitante</b></td>
					<td colspan="2"><%=nm_solicitante%></td>
					<td colspan="2">&nbsp;</td>							
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Unidade</b></td>
					<td colspan="4"><%=nm_unidade%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">	
					<td  align="center"><b>Qtd</b></td>
					<td colspan="2"><%=num_qtd%></td>
					<td align="center"><b>Especialidade</b></td>
					<td colspan="2"><%=nm_especialidade%>
					</td>
											
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Equip./instr.</b></td>
					<td colspan="2"><%=nm_equipamento%></td>
					<td  align="center"><b>Marca</b></td>
					<td><%=nm_marca%></td>	
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>N/S</b></td>
					<td colspan="4"><%=cd_ns%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Patrim�mio</b></td>
					<td colspan="4"><%=no_patrimonio%> <span style="color:silver;">(<%=no_pat_old%>)</span></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Motivo</b></td>
					<td colspan="4"><%=motivo%></td>
				</tr>
			<!--Fim da etapa 1.1-->			
			<!--In�cio da etapa 1.2-->				
				<%if cd_user <> usuario_restrito then%>
				<tr style="visibility:<%=visual%>;">
					<td colspan="5" align="right"><!--Separador de etapas-->&nbsp;
						<%if jan = 0 then%>
							<a href="../manutencao2.asp?tipo=nova2&os_alt=1&cd_unidade2=<%=cd_unidade%>&num_os2=<%=num_os%>&jan=<%=jan%>"> alterar &nbsp; </a>
						<%elseif jan = 1 then%>
							<a href="manutencao_nova_3.asp?tipo=nova2&cd_unidade2=<%=cd_unidade%>&cd_codigo=<%=cd_codigo%>&campo=cd_codigo&jan=<%=jan%>"></a>
							<a href="manutencao_nova.asp?tipo=nova2&os_alt=1&cd_unidade2=<%=cd_unidade%>&num_os2=<%=num_os%>">alterar &nbsp; </a>
						<%end if%></a>
					</td>
				</tr>
			</table>
			<br>
			<table border="1" style="<%=css_tabela%>">	
				<%end if%>	
		<%if cd_user <> usuario_restrito then%>
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">PROVID�NCIA</td></tr>
				<tr bgcolor="#<%=bg_cor%>">
				  	<td  align="center"><b>Data envio</b></td>
					<td><input type="text" name="dt_entrada" size="11" maxlength="50" class="inputs" value="<%=dt_entrada%>" onFocus="nextfield ='cd_avaliacao';"></td>
					<td  align="center"><b>Avalia��o</b></td>
					<td colspan="2"><select name="cd_avaliacao" class="inputs" onFocus="nextfield ='cd_fornecedor';">
						<option value="">Selecione</option>
						<%
						'****** Mostra as op��es ***
						selecao = " * "
						tabela = " TBL_avaliacao"
						condicoes = ""
						criterios = " Order by nm_avaliacao"

						strsql ="up_listagem @selecao='"&selecao&"', @tabela='"&tabela&"', @condicoes='"&condicoes&"', @criterios='"&criterios&"'"
						Set rs_aval = dbconn.execute(strsql)%>
						
						<%
						Do while not rs_aval.EOF
						cd_aval = rs_aval("cd_codigo")
						nm_avaliacao = rs_aval("nm_avaliacao")
						
						'if cd_unidade = cd_aval Then
						if cd_avaliacao = cd_aval Then
							ck_aval = "selected"
						'Elseif int(cd_aval) = 2 Then
						'	ck_aval = "selected"
						end if%>
						<option value="<%=cd_aval%>" <%=ck_aval%>><%=nm_avaliacao%></option>
						<%rs_aval.movenext
						ck_aval = ""
						loop
						%>
					 
					</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>"><td  align="center"><b>Encaminhar para</b></td>
						<%strsql ="TBL_Fornecedor where cd_status = 1 order by nm_fornecedor"
					  	Set rs_forn = dbconn.execute(strsql)%>
					<td colspan="4"><span id='divForn'><select name="cd_fornecedor" class="inputs" onFocus="nextfield ='observacoes';">
						<option value="">Selecione</option>
						<%Do while Not rs_forn.EOF
							cod_forn = rs_forn("cd_codigo")
							fornecedor = rs_forn("nm_fornecedor")
							if cd_fornecedor = cod_forn Then
								ck_forn="selected"
							'elseif cod_forn = 28 Then
							'	ck_forn="selected"
							else 
								ck_forn=""
							end if%>
						<option value="<%=rs_forn("cd_codigo")%>" <%=ck_forn%>><%=fornecedor%></option>
						<%rs_forn.movenext
					 	Loop%>
						</select>
						</span>
						<img src="../imagens/reload6.gif" type="button" alt="Atualizar" border="0" onClick="fill_forn();">
						<a href="javascript: void(0);" onclick="window.open('http://smartnew:89/ferramentas/fornecedores_cad.asp', 'principal', 'width=530, height=270'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a>
						
						</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>"><td  align="center"><b>Cadastrado por</b></td>
					<td colspan="4"><%=cd_liberacao%> <b>em:</b> <%=dt_cad%></td>
					
				</tr>	
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Ocorr�ncia</b><br></td>
					<td colspan="4">
						<textarea cols="81" rows="2" name="observacoes" class="inputs" onFocus="nextfield ='dt_manut_orc';"><%=observacoes%></textarea></td>
				</tr>				
				<tr><td colspan="5">&nbsp;</td></tr>
				<!--Os novos campos come�am aqui-->
		<%end if%>
				<%
				sequencia = "2"
				xsql_andamento = "SELECT * FROM View_andamento_lista2 Where cd_cod = "&cd_cod_andamento&" ORDER BY cd_os"
				Set rs_andamento = dbconn.execute(xsql_andamento)
					
				If rs_andamento.EOF Then
				'Response.write("fim")
				Else 
				'response.write("ok")
				
				cd_cod = rs_andamento("cd_codigo")
				'cod = rs_andamento("cd_cod")
				dt_manut_orc = rs_andamento("dt_manut_orc")
				nm_fornecedor = rs_andamento("nm_fornecedor")
				nm_contato = rs_andamento("contato")
				dt_prev_resposta = rs_andamento("dt_prev_resposta")
				dt_resposta_orc = rs_andamento("dt_resposta_orc")
				cd_orcamento = rs_andamento("cd_orcamento")
				nm_orcamento = rs_andamento("nm_orcamento")
				num_os_assist = rs_andamento("num_os_assist")
				dt_prev_retorno = rs_andamento("dt_prev_retorno")
				cd_situacao = rs_andamento("cd_situacao")
				cd_alerta = rs_andamento("cd_alerta")
				cd_qtd_valor = rs_andamento("cd_qtd_valor")
				cd_valor_unit = rs_andamento("cd_valor_unit")
				
				
				dt_prev_retorno = zero(day(dt_prev_retorno))&"/"&zero(month(dt_prev_retorno))&"/"&right(year(dt_prev_retorno),2)
				
				
				
				
				
					if not cd_valor_unit = "" Then
						if instr(cd_valor_orc,".") = 1 then
							cd_valor_unit = abs(replace(cd_valor_unit,".",","))
						end if
						'cd_valor_unit = formatnumber(cd_valor_unit)
					end if
				cd_valor_orc = rs_andamento("cd_valor_orc")
					if not cd_valor_orc = "" Then
						if instr(cd_valor_orc,".") = 1 then
							cd_valor_orc = abs(replace(cd_valor_orc,".",","))
							
							
							
						end if
						'cd_valor_orc = formatnumber(cd_valor_orc)
					end if
				comentarios = rs_andamento("comentarios")
				dt_retorno = rs_andamento("dt_retorno")
				dt_retorno_un = rs_andamento("dt_retorno_un")
				nm_natureza_defeito = rs_andamento("nm_natureza_defeito")
				cd_gestao = rs_andamento("cd_gestao")
				sequencia = "2"
				
				
				strsql = "SELECT TOP (200) nm_valor FROM TBL_manutencao_orcamento WHERE (cd_codigo = '"&cd_orcamento&"')"
				Set rs2 = dbconn.execute(strsql)
					if not rs2.EOF Then
						valor_total_orcamento = formatnumber(rs2("nm_valor"),2)
					end if
				
				End if
				
				if cd_cod = "" Then
				acao = 1
				Else 
				acao = 2
				End if
				%>
				<input type="hidden" name="cd_os" value="<%=num_os%>">
				<input type="hidden" name="cd_cod" value="<%=cd_codigo%>">
				<input type="hidden" name="cd_codigo_os" value="<%=cd_cod_andamento%>">
				<input type="hidden" name="sequencia" value="<%=int(sequencia)%>">
				<input type="hidden" name="cd_natureza_os" value="<%=cd_natureza_os%>">
				
				<input type="hidden" name="cd_orcamento" id="codOrc" size="11" value="<%=cd_orcamento%>">
				<input type="hidden" name="dt_manut_orc" id="dtOrca" size="11" value="<%=dt_manut_orc%>">
				
				<%if cd_user <> usuario_restrito then%>
				<%if esconder = "0" Then
				Else%>
				
				<%
				if dt_resposta_orc <> "" then
				
				dt_retorno = zero(day(dt_retorno))&"/"&zero(month(dt_retorno))&"/"&right(year(dt_retorno),2)
				dt_resposta_orc = zero(day(dt_resposta_orc))&"/"&zero(month(dt_resposta_orc))&"/"&right(year(dt_resposta_orc),2)
				dt_cad = zero(day(dt_cad))&"/"&zero(month(dt_cad))&"/"&right(year(dt_cad),2)
				
				dt_retorno_un = zero(day(dt_retorno_un))&"/"&zero(month(dt_retorno_un))&"/"&right(year(dt_retorno_un),2)
				
				
				
				
				end if
				
				%>
				<!--tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">MANUTEN��O / OR�AMENTO <b style="color:#b4b4b4;"><%=cd_cod_andamento%></b></td>
				</tr-->
				<!--tr bgcolor="#<%=bg_cor%>">
					<td style="color:red;" align="center">** ALERTA **</td>
					<td><input type="checkbox" name="cd_alerta" value="1" <%if cd_alerta = 1 Then%>checked<%end if%>></td>
					<td colspan="3">&nbsp;</td>
				</tr-->
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Ass. T�c / Fornecedor</td>
					<!--input type="hidden" name="cd_fornecedor" value="<%=cd_fornecedor%>"-->
					<td ><%=nm_fornecedor%><%'=cod%></td>
					<td align="center">Contato</td>
					<td><input type="text" name="nm_contato" size="15" class="inputs"  value="<%=nm_contato%>"></td>
					<td>&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Previs�o resposta</td>
					<td><input type="text" name="dt_prev_resposta" size="11" maxlength="20" class="inputs" value="<%=dt_prev_resposta%>"></td>
					<td colspan="1">&nbsp;</td>
				</tr>
					<tr bgcolor="#<%=bg_cor%>">					
					<td align="center">Motivo</td>
					<td  colspan="2">
					<select name="nm_natureza_defeito" class="inputs">
						<option value=""></option>
						<option value="1" <%if nm_natureza_defeito = 1 then response.write("selected") end if%>>Motivo desconhecido</option>
						<option value="2" <%if nm_natureza_defeito = 2 then response.write("selected") end if%>>Falha do operador</option>
						<option value="3" <%if nm_natureza_defeito = 3 then response.write("selected") end if%>>Desgaste por uso</option>
						<option value="4" <%if nm_natureza_defeito = 4 then response.write("selected") end if%>>Quebra por terceiros</option>
						<option value="5" <%if nm_natureza_defeito = 5 then response.write("selected") end if%>>Solicita��o de compras por terceiros</option>						
					</select>
					</td>
					<td colspan="2">&nbsp;</td>
				</tr>
			</table>
			<br>
			<table  border="1" style="<%=css_tabela%>">
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">OR�AMENTO <b style="color:#b4b4b4;"><%=cd_cod_andamento%></b></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">N� Or�.</td>
					<td colspan="1"><input type="text" name="num_os_assist" size="15" maxlength="100" class="inputs" value="<%=num_os_assist%>" onkeyup="mostra_orcamento2(this.value,<%=cd_fornecedor%>);"><%if cd_user = 46 then response.write("("&cd_orcamento&")")%></td>
					<td align="center">Data</td>
					<td colspan="2"><span id="divorcamento" style="font-size:11px;"><b><%=dt_manut_orc%></b></span></td>
					
				</tr>
				
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center">Data aprov.<br> or�amento</td>
					<td><input type="text" name="dt_resposta_orc" id="dtAprov" size="11" maxlength="20" class="inputs" value="<%=dt_resposta_orc%>"></td>
					<td align="center">Valor or�.</td>
					<td align="right"><b><%=valor_total_orcamento%></b></td></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">	
					<td  align="center">Or�amento</td>
					<%if nm_orcamento = "0" Then%>
					<%ck_mo10 = "selected"
					'response.write("zero")%>
					<%elseif nm_orcamento = "1" Then%>
					<%ck_mo11 = "selected"
					'response.write("um")%>
					<%elseif nm_orcamento = "2" Then%>
					<%ck_mo12 = "selected"
					'response.write("dois")%>
					<%elseif nm_orcamento = "3" Then%>
					<%ck_mo13 = "selected"
					'response.write("tr�s")%>
					<%End if%>
					<td><select name="nm_orcamento" class="inputs">
							<option value="" <%=ck_mo10%>>&nbsp;</option>
							<option value="4" <%=ck_mo14%>>Reprovado</option>
							<option value="1" <%=ck_mo11%>>Aguardando</option>
							<option value="2" <%=ck_mo12%>>Aprovado</option>
							<option value="3" <%=ck_mo13%>>Garantia</option>
						</select>
					</td>
				</tr>
				</tr>
				<tr><td colspan="4">&nbsp;</td></tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center">R$ (unit�rio)</td>
					<td colspan="2"><input type="text" class="inputs" name="cd_valor_unit" size="10" maxlength="10" value="<%if isnumeric(cd_valor_unit) Then response.write(formatnumber(cd_valor_unit,2))%>" onKeyup="valores();" style="text-align:right"> &nbsp; &nbsp; &nbsp; 
					Unidades <input type="text" class="inputs" name="cd_qtd_valor" size="5" maxlength="5" onKeydown="valores();" value="<%=cd_qtd_valor%>" onKeyup="valores();"> </td>
					<td  align="center"><b>Valor Total R$</b></td>
					<td><input type="text" name="cd_valor_orc" readonly="1" size="11" maxlength="20" class="inputs" value="<%if isnumeric(cd_valor_orc) Then response.write(formatnumber(cd_valor_orc,2))%>" onKeyup="valores();" style="text-align:right" ></td>
				</tr>
				<tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Situa��o</td>
						<%strsql ="SELECT * From TBL_situacao Order by nm_situacao"
					  	Set rs_situac = dbconn.execute(strsql)%>
				
						<td colspan="2">
						<span id='divSit'><select name="cd_situacao" class="inputs">
										<option value=""></option>
						<%Do while Not rs_situac.EOF%><%strsituacao = rs_situac("cd_codigo")%>
						
						<%if cd_situacao = strsituacao then%><%ck_sit = "selected"%><%end if%>
										<option value="<%=rs_situac("cd_codigo")%>" <%=ck_sit%>><%=rs_situac("nm_situacao")%></option>	
										
						<%ck_sit = ""%>
						<%rs_situac.movenext%>
						<%Loop%>
						</select>
						</span>
						<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_sit();">
						<a href="javascript: void(0);" onclick="window.open('janelas_cadastros/situacao_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=370, height=165'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->					
					</td>
					<td  align="center">Previs�o retorno/entrega</td>
					<td><input type="text" name="dt_prev_retorno" size="11" maxlength="20" class="inputs" value="<%=dt_prev_retorno%>"></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center">Laudo de <br>manuten��o</td>
					<td colspan="2">
						<textarea cols="40" rows="2" name="comentarios" class="inputs"><%=comentarios%></textarea></td>
				</tr>
				<tr>
					<td  align="center">Gest�o</td>
						<%strsql ="SELECT * From TBL_gestao Order by cd_ordem"
					  	Set rs_gestao = dbconn.execute(strsql)%>
				
						<td colspan="1">
							<select name="cd_gestao" class="inputs">
								<option value="1"></option>
								<%Do while Not rs_gestao.EOF%><%strgestao = rs_gestao("cd_codigo")%>
								<%if cd_gestao = strgestao then%><%ck_gest = "selected"%><%end if%>
									<option value="<%=rs_gestao("cd_codigo")%>" <%=ck_gest%>><%=rs_gestao("nm_gestao")%></option>
								<%ck_gest = ""%>
								<%rs_gestao.movenext%>
								<%Loop%>
						</select>
					</td>
					
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Data de entrega/retorno</td>
					<td colspan="2"><input type="text" name="dt_retorno" size="11" maxlength="20" class="inputs" value="<%=dt_retorno%>"></td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Data de envio para Hospital</td>
					<td colspan="2"><input type="text" name="dt_retorno_un" size="11" maxlength="20" class="inputs" value="<%=dt_retorno_un%>"></td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr><td colspan="10"><!-- *** Inclui o registro de ocorrencias *** --><%'=cd_unidade&"."&manutencao(num_os)%>
				<%'id_origem = num_os
				id_origem = zero(cd_unidade)&"."&manutencao(num_os)
				cd_origem = 3
				pag_retorno = ".."&mem_posicao
				variaveis = "../../manutencao2.asp?tipo=andamento$cd_codigo="&cd_cod_andamento&"$campo=cd_codigo"%>
				<!--#include file="../ocorrencia/ocorrencia_formulario.asp"-->
				</td></tr>				
				<%end if%>
			</table>
			<br>
			<table border="1" style="<%=css_tabela%>">	
				<tr style="visibility:<%=visual%>;">
					<td colspan="100%" align="center">
						<input type="hidden" name="cd_codigo" value="<%=cd_cod%>">
						<input type="hidden" name="acao" value="<%=acao%>">
						<input type="submit" value="Enviar"><%'=acao%>						
					</td>
				</tr>
				
				
				
				
				<%end if%>
				
				
				
			<!--Fim da etapa 1.2-->
			
						</form>
			<%end if%>				
			</table>
		</td>
	</tr>
</table>
<br>

<script type="text/javascript">
	formataMoeda(document.forms.form.cd_valor_unit, 2);
	formataMoeda(document.forms.form.cd_valor_orc, 2);

function JsDelete(cod1, cod2, cod3)
{
  if (confirm ("Tem certeza que deseja excluir a ocorr�ncia?"))
	  {
		document.location.href('acoes/manutencao_acao.asp?cd_cod_ocorrencia='+cod1+'&cd_codigo='+cod2+'&variaveis='+cod3+'');
	  }
}
</SCRIPT>


</BODY>
</HTML>



