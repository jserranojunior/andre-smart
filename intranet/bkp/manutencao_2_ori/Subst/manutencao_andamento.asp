<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/estilos_css.htm"--> 
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>

<style type="text/css">
<!--
#no_print{ 
	visibility:visible; 
	display: block;}
#ok_print{
	visibility:hidden; 
	display: none;}
-->
</style>
<%esconder = "1"%>
<BODY>
<br> 
<table align="center" border="1" cellspacing="0" cellpadding="0">
	<tr>
		<td>
 <SCRIPT LANGUAGE="javascript">
    {       
       objVld=getValidator('form');       
      
	   objVld.addRule('num_os',['required','texto'],'"O.S."');
	   objVld.addRule('dt_os',['required','texto'],'"Data da solicitação"');
	   objVld.addRule('nm_solicitante',['required','texto'],'"Solicitante"');
	   objVld.addRule('nm_unidade',['required','texto'],'"Unidade"');
   	   objVld.addRule('num_qtd',['required','texto'],'"Qtd"');
	   objVld.addRule('cd_equipamento',['required','texto'],'"Equip/Instr"');
	   objVld.addRule('cd_marca',['required','texto'],'"Marca"');
	   objVld.addRule('motivo',['required','texto'],'"Motivo"');
	   objVld.addRule('dt_entrada',['required','texto'],'"Data da provicência"');
	   objVld.addRule('cd_avaliacao',['required','texto'],'"Avaliação"');
	   objVld.addRule('nm_unidade',['required','texto'],'"Unidade"');
   }
</SCRIPT>
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
	var	valor = valor.replace('.','');
	var	valor = valor.replace(',','.');	
	var valor_total = parseFloat(valor*qtd);
		document.form.cd_valor_orc.value = valor_total.toFixed(2);
	


	
	
	
	
	}

// End -->
</script>
<!--#include file="../js/ajax.js"-->
<!--#include file="../js/manutencao.js"-->
<%
session_cd_usuario = session("cd_codigo")
usuario_restrito = int("57")

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

xsql ="up_os_lista @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
Set rs = dbconn.execute(xsql)
If rs.EOF Then%>


			<table width="550" border="1" cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
				<tr>		
					<td class="txt_cinza" colspan="6">
					<b>Manutenção &raquo; - <font color="red">Listagem1</font></b><br><br></td>
					<!--td class="txt_cinza" align="right" valign="top">Busca O.S.: &nbsp;</td>
					<td class="txt_cinza" colspan="2">
					<form action="ordemservico_andamento.asp" name="busca" id="busca">
					<input type="text" name="cd_codigo" size="6" maxlength="6">
					<input type="hidden" name="campo" value="num_os">
					<input type="hidden" name="ordem" value="1">		
					<input type="submit" value="ok">
					</form-->
					</td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao.asp" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova">
					<tr id="no_print"><td colspan="100%">NOVA O.S.:  <input type="text" name="cd_unidade" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"><input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%if session_cd_usuario <> usuario_restrito then%>
				<tr style="visibility:<%=visual%>;">
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao.asp?tipo=pendencias&filtro=pendentes">Pendentes</a> 
						/ <a href="manutencao.asp?tipo=pendencias&filtro=encerradas&strtop=20">Encerradas</a>
						/ <a href="manutencao.asp?tipo=pendencias&filtro=todas&strtop=20">Todas2</a>						
					</td>
				</tr>
				<tr><td><br>*** Não foi possível encontrar a O.S. solicitada ***<br>&nbsp;</td></tr>
				<%end if
			registros = 0
			Else
			Do while not rs.EOF
				cd_cod_andamento = rs("cd_codigo")
				cd_unidade = rs("cd_unidade")
				num_os = rs("num_os")
				dt_os = rs("dt_os")
				nm_unidade = rs("nm_unidade")
				nm_especialidade = rs("nm_especialidade")
				nm_equipamento = rs("nm_equipamento")
			
				dt_entrada = rs("dt_entrada")
				cd_avaliacao = rs("cd_avaliacao")
				nm_fornecedor = rs("nm_fornecedor")
				cd_fornecedor = rs("cd_fornecedor")
				cd_liberacao = rs("cd_liberacao")
			
				cd_codigo = rs("cd_codigo")
			
				num_qtd = rs("num_qtd")
				nm_marca = rs("nm_marca")
				cd_ns = rs("cd_ns")
				cd_patrimonio = rs("cd_patrimonio")
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
			%>
			<table width="550" border="1" cellspacing="1" cellpadding="0" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
				<tr>		
					<td class="txt_cinza" colspan="6">
					<b>Manutenção &raquo; - <font color="red">Listagem</font></b><b style="color=white;"><%=session_cd_usuario%> <%=usuario_restrito%></b><br><br></td>
					<!--td class="txt_cinza" align="right" valign="top">Busca O.S.: &nbsp;</td>
					<td class="txt_cinza" colspan="2">
					<form action="ordemservico_andamento.asp" name="busca" id="busca">
					<input type="text" name="cd_codigo" size="6" maxlength="6">
					<input type="hidden" name="campo" value="num_os">
					<input type="hidden" name="ordem" value="1">		
					<input type="submit" value="ok">
					</form-->
					</td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao.asp?tipo=nova" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova">
					<tr id="ok_print"><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<%if session_cd_usuario <> usuario_restrito then%>
				<tr style="visibility:<%=visual%>;">
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao.asp?tipo=pendencias&filtro=pendentes">Pendentes</a> 
						/ <a href="manutencao.asp?tipo=pendencias&filtro=encerradas&strtop=20">Encerradas</a>
						/ <a href="manutencao.asp?tipo=pendencias&filtro=todas&strtop=20">Todas</a>						
					</td>
				</tr>
				<%end if%>				
			<!--Início da etapa 1.1-->
			<%jan=request("jan")
			if jan = 1 then
			gambi = "../"
			end if%>
				<tr>			
					<form method="post" action="<%=gambi%>manutencao/acoes/manutencao_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="filtro" value="andamento">
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
				<tr bgcolor="#b4b4b4">
					<td  colspan="100%" align="center" class="textopadrao">SOLICITAÇÃO </td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>O.S.</b></td>
					<td colspan="4"><%=zero(cd_unidade)%>.<%=manutencao(num_os)%></td>
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
					<td align="center"><b>Patrimômio</b></td>
					<td colspan="4"><%=cd_patrimonio%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Motivo</b></td>
					<td colspan="4"><%=motivo%></td>
				</tr>
			<!--Fim da etapa 1.1-->			
			<!--Início da etapa 1.2-->				
				<%if session_cd_usuario <> usuario_restrito then%>
				<tr style="visibility:<%=visual%>;"><td colspan="100%" align="right"><!--Separador de etapas-->&nbsp;<a href="manutencao.asp?tipo=nova&cd_unidade=<%=cd_unidade%>&cd_codigo=<%=cd_codigo%>&campo=cd_codigo"> alterar</a></td></tr>
				<tr><td colspan="5">&nbsp;</td></tr>
				<%end if%>	
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">PROVIDÊNCIA</td></tr>
				<tr bgcolor="#<%=bg_cor%>">
				  	<td  align="center"><b>Data envio</b></td>
					<td><input type="text" name="dt_entrada" size="11" maxlength="50" class="inputs" value="<%=dt_entrada%>" onFocus="nextfield ='cd_avaliacao';"></td>
					<td  align="center"><b>Avaliação</b></td>
					<td colspan="2"><select name="cd_avaliacao" class="inputs" onFocus="nextfield ='cd_fornecedor';">
						<option value="">Selecione</option>
						<%
						'****** Mostra as opções ***
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
						end if%>
						<option value="<%=cd_aval%>" <%=ck_aval%>><%=nm_avaliacao%></option>
						<%rs_aval.movenext
						ck_aval = ""
						loop
						%>
					 
					</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>"><td  align="center"><b>Encaminhar para</b></td>
						<%strsql ="TBL_Fornecedor order by nm_fornecedor"
					  	Set rs_forn = dbconn.execute(strsql)%>
					<td colspan="4"><span id='divForn'><select name="cd_fornecedor" class="inputs" onFocus="nextfield ='observacoes';">
						<option value="">
						<%Do while Not rs_forn.EOF%><%fornecedor=rs_forn("nm_fornecedor")%><%if nm_fornecedor=fornecedor Then%><%ck_forn="selected"%><%else ck_forn=""%><%end if%>
						<option value="<%=rs_forn("cd_codigo")%>" <%=ck_forn%>><%=rs_forn("nm_fornecedor")%></option><%rs_forn.movenext
					 	Loop%>
						</select>
						</span>
						<img src="../imagens/reload6.gif" type="button" alt="Atualizar" border="0" onClick="fill_forn();">
						<a href="javascript: void(0);" onclick="window.open('janelas_cadastros/fornecedor_cad.asp?cd_cod_horas=<%=cd_cod_horas%>&cd_cod_abono=<%=cd_cod_abono%>&janela=pop_up&metodo=fecha', 'principal', 'width=530, height=270'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a>
						
						</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>"><td  align="center"><b>Cadastrado por</b></td>
					<td colspan="4"><%=cd_liberacao%><%'=nm_nome%><%'=session("nm_nome")%></td>
				</tr>	
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Ocorrência</b><br></td>
					<td colspan="4">
						<textarea cols="81" rows="2" name="observacoes" class="inputs" onFocus="nextfield ='dt_manut_orc';"><%=observacoes%></textarea></td>
				</tr>				
				<tr><td colspan="5">&nbsp;</td></tr>
				<!--Os novos campos começam aqui-->
				<%
				sequencia = "2"
				'xsql_andamento ="up_andamento_lista @cd_unidade="&cd_unidade&", @cd_codigo='"&num_os&"',@campo=cd_os, @sequencia='"&sequencia&"'"
				
				'xsql_andamento ="up_andamento_lista @cd_unidade='', @cd_codigo='"&num_os&"',@campo=cd_os, @sequencia='"&sequencia&"'"
				'xsql_andamento = "SELECT * FROM View_andamento_lista Where cd_cod = "&cd_cod_andamento&" ORDER BY cd_os"
				'xsql_andamento = "SELECT * FROM View_andamento_lista Where cd_cod = "&cd_cod_andamento&" ORDER BY cd_os"
				xsql_andamento = "SELECT * FROM View_andamento_lista Where cd_cod = "&cd_cod_andamento&" ORDER BY cd_os"
				Set rs_andamento = dbconn.execute(xsql_andamento)
					
				If rs_andamento.EOF Then
				'Response.write("fim")
				Else 
				'response.write("ok")
				
				cd_cod = rs_andamento("cd_codigo")
				'cod = rs_andamento("cd_cod")
				dt_manut_orc = rs_andamento("dt_manut_orc")
				nm_fornecedor = rs_andamento("nm_fornecedor")
				nm_contato = rs_andamento("nm_contato")
				dt_prev_resposta = rs_andamento("dt_prev_resposta")
				dt_resposta_orc = rs_andamento("dt_resposta_orc")
				nm_orcamento = rs_andamento("nm_orcamento")
				num_os_assist = rs_andamento("num_os_assist")
				dt_prev_retorno = rs_andamento("dt_prev_retorno")
				cd_situacao = rs_andamento("cd_situacao")
				cd_alerta = rs_andamento("cd_alerta")
				cd_qtd_valor = rs_andamento("cd_qtd_valor")
				cd_valor_unit = rs_andamento("cd_valor_unit")
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
				sequencia = "2"
				
				If dt_manut_orc = "1/1/1950" OR dt_manut_orc = "01/01/1950" Then
				dt_manut_orc = ""
				End if
				If dt_resposta_orc = "1/1/1950" OR dt_resposta_orc = "01/01/1950" Then
				dt_resposta_orc = ""
				End if
				If dt_prev_resposta = "1/1/1950" OR dt_prev_resposta = "01/01/1950" Then
				dt_prev_resposta = ""
				End if
				If dt_prev_retorno = "1/1/1950" OR dt_prev_retorno = "01/01/1950" Then
				dt_prev_retorno = ""
				End if
				If dt_retorno = "1/1/1950" OR dt_retorno = "01/01/1950" Then
				dt_retorno = ""
				End if
				If dt_retorno_un = "1/1/1950" OR dt_retorno_un = "01/01/1950" Then
				dt_retorno_un = ""
				End if
				
				
				
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
				<%if session_cd_usuario <> usuario_restrito then%>
				<%if esconder = "0" Then
				Else%>
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">MANUTENÇÃO / ORÇAMENTO <b style="color:#b4b4b4;"><%=cd_cod_andamento%></b></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td style="color:red;" align="center">** ALERTA **</td>
					<td><input type="checkbox" name="cd_alerta" value="1" <%if cd_alerta = 1 Then%>checked<%end if%>></td>
					<td colspan="3">&nbsp;</td>
				</tr>	
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center">Data manut. / orc.</td>
					<td><input type="text" name="dt_manut_orc" size="11" maxlength="50" class="inputs" value="<%=dt_manut_orc%>"></td> 
					<td  align="center">Assist. Técnica / Fornecedor</td>
					<!--input type="hidden" name="cd_fornecedor" value="<%=cd_fornecedor%>"-->
					<td colspan="2"><%=nm_fornecedor%><%'=cod%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					
					<td align="center">Contato</td><td><input type="text" name="nm_contato" size="15" class="inputs"  value="<%=nm_contato%>"></td>
					<td  align="center">Previsão resposta</td>
					<td><input type="text" name="dt_prev_resposta" size="11" maxlength="20" class="inputs" value="<%=dt_prev_resposta%>"></td>
					<td>&nbsp;</td>
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
						<option value="5" <%if nm_natureza_defeito = 5 then response.write("selected") end if%>>Solicitação de compras por terceiros</option>						
					</select>
					</td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Data da resposta do orçamento</td>
					<td colspan="2"><input type="text" name="dt_resposta_orc" size="11" maxlength="20" class="inputs" value="<%=dt_resposta_orc%>"></td>
					<td  align="center">Orçamento</td>
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
					'response.write("três")%>
					<%End if%>
					
					<td><select name="nm_orcamento" class="inputs">
																<option value="0" <%=ck_mo10%>>Reprovado</option>
																<option value="1" <%=ck_mo11%>>Aguardando</option>
																<option value="2" <%=ck_mo12%>>Aprovado</option>
																<option value="3" <%=ck_mo13%>>Garantia</option>
					</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">O.S. Assist. Téc.</td>
					<td colspan="2"><input type="text" name="num_os_assist" size="35" maxlength="100" class="inputs" value="<%=num_os_assist%>"></td>
					<td  align="center">Valor R$</td>
					<td><input type="text" name="cd_valor_orc" size="11" maxlength="20" class="inputs" value="<%=cd_valor_orc%>"></td>
					
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center">Unidade</td>
					<td colspan="2"><input type="text" name="cd_qtd_valor" size="5" maxlength="5" onKeydown="valores();" value="<%=cd_qtd_valor%>"> &nbsp; &nbsp; &nbsp; Valor unitário R$ <input type="text" name="cd_valor_unit" size="10" maxlength="10" value="<%=cd_valor_unit%>" onKeyup="valores();"></td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center">Situação</td>
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
						<img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_sit();">
						<a href="javascript: void(0);" onclick="window.open('janelas_cadastros/situacao_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=370, height=165'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a>					
					</td>
					<td  align="center">Previsão retorno/entrega</td>
					<td><input type="text" name="dt_prev_retorno" size="11" maxlength="20" class="inputs" value="<%=dt_prev_retorno%>"></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center">Laudo de <br>manutenção</td>
					<td colspan="2">
						<textarea cols="40" rows="2" name="comentarios" class="inputs"><%=comentarios%></textarea></td>
				</tr>
				<tr>
					<td  align="center">Gestão</td>
						<%strsql ="SELECT * From TBL_gestao Order by cd_ordem"
					  	Set rs_gestao = dbconn.execute(strsql)%>
				
						<td colspan="1">
							<select name="cd_gestao" class="inputs">
								<option value=""></option>
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
				<tr><td colspan="100%"><!-- *** Inclui o registro de ocorrencias *** --><%'=cd_unidade&"."&manutencao(num_os)%>
				<%'id_origem = num_os
				id_origem = zero(cd_unidade)&"."&manutencao(num_os)
				cd_origem = 3
				pag_retorno = ".."&mem_posicao
				variaveis = "../../manutencao.asp?tipo=andamento$cd_codigo="&cd_cod_andamento&"$campo=cd_codigo"%>
				<!--#include file="../ocorrencia/ocorrencia_formulario.asp"-->
				</td></tr>				
				<%end if%>
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

<SCRIPT LANGUAGE="javascript">
 {
        MaskInput(document.form.dt_manut_orc, "99/99/9999");
	    MaskInput(document.form.dt_prev_resposta, "99/99/9999");
		MaskInput(document.form.dt_resposta_orc, "99/99/9999");
		MaskInput(document.form.dt_prev_retorno, "99/99/9999");
		MaskInput(document.form.dt_retorno, "99/99/9999");
		MaskInput(document.form.dt_retorno_un, "99/99/9999");
 }
</SCRIPT>
<SCRIPT LANGUAGE="javascript">

function JsDelete(cod1, cod2, cod3)
{
  if (confirm ("Tem certeza que deseja excluir a ocorrência?"))
	  {
		document.location.href('acoes/manutencao_acao.asp?cd_cod_ocorrencia='+cod1+'&cd_codigo='+cod2+'&variaveis='+cod3+'');
	  }
}
</SCRIPT>


</BODY>
</HTML>



