<!--#include file="../includes/inc_open_connection.asp"--> 
<!--#include file="../includes/estilos_css.htm"--> 
<script language="JavaScript" type="text/javascript" src="js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="js/mascarainput.js"></SCRIPT>
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
	   objVld.addRule('dt_os',['required','texto'],'"Data da solicita��o"');
	   objVld.addRule('nm_solicitante',['required','texto'],'"Solicitante"');
	   objVld.addRule('nm_unidade',['required','texto'],'"Unidade"');
   	   objVld.addRule('num_qtd',['required','texto'],'"Qtd"');
	   objVld.addRule('cd_equipamento',['required','texto'],'"Equip/Instr"');
	   objVld.addRule('cd_marca',['required','texto'],'"Marca"');
	   objVld.addRule('motivo',['required','texto'],'"Motivo"');
	   objVld.addRule('dt_entrada',['required','texto'],'"Data da provic�ncia"');
	   objVld.addRule('cd_avaliacao',['required','texto'],'"Avalia��o"');
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
// End -->
</script>
<%
cd_codigo = request("cd_codigo")
num_os = request("num_os")
campo = request("campo")
ordem = request("ordem")
list = request("list")
lista = "1"
strvisual = request("visual")


if strvisual = "1" Then
visual = "hidden"
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
					<b>Manuten��o &raquo; - <font color="red">Listagem1</font></b><br><br></td>
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
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manut_os_nova.asp" name="busca" id="busca">
					<tr><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<tr style="visibility:<%=visual%>;">
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao.asp?strtop=20&tipo=pendentes">Pendentes</a> 
						/ <a href="manutencao.asp?strtop=20&tipo=encerradas">Encerradas</a>
						/ <a href="manutencao.asp?strtop=20&tipo=todas">Todas</a>
					</td>
				</tr>
				<tr><td><br>*** N�o foi poss�vel encontrar a O.S. solicitada ***<br>&nbsp;</td></tr>
				<%
			registros = 0
			Else
			Do while not rs.EOF
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
					<b>Manuten��o &raquo; - <font color="red">Listagem</font></b><br><br></td>
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
				<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manut_os_nova.asp" name="busca" id="busca">
					<tr><td colspan="100%">NOVA O.S.:<input type="text" name="num_os" size="10" maxlength="10"><input type="submit" value="ok"></td></tr>
					</form>
					<tr><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<tr style="visibility:<%=visual%>;">
				    <td colspan="7" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao.asp?strtop=20&tipo=pendentes">Pendentes</a> 
						/ <a href="manutencao.asp?strtop=20&tipo=encerradas">Encerradas</a>
						/ <a href="manutencao.asp?strtop=20&tipo=todas">Todas</a>
					</td>
				</tr>					
			<!--In�cio da etapa 1.1-->
				<tr>			
					<form method="post" action="manutencao/manut_os_acao.asp" name="form"  id="form" onsubmit="return validar_cad(document.form);">
						<input type="hidden" name="tipo" value="andamento">
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
				<tr bgcolor="#b4b4b4">
					<td  colspan="100%" align="center" class="textopadrao">SOLICITA��O</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center"><b>O.S.</b></td>
					<td colspan="4"><%=num_os%></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center"><b>Data</b></td>
					<td><%=dt_os%></td>
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center"><b>Solicitante</b></td>
					<td colspan="2"><%=nm_solicitante%></td>
					<td colspan="2">&nbsp;</td>							
				<tr bgcolor="#f5f5f5">
					<td  align="center"><b>Unidade</b></td>
					<td colspan="4"><%=nm_unidade%></td>
				</tr>
				<tr bgcolor="#f5f5f5">	
					<td  align="center"><b>Qtd</b></td>
					<td colspan="2"><%=num_qtd%></td>
					<td align="center"><b>Especialidade</b></td>
					<td colspan="2"><%=nm_especialidade%>
					</td>
											
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center"><b>Equip./instr.</b></td>
					<td colspan="2"><%=nm_equipamento%></td>
					<td  align="center"><b>Marca</b></td>
					<td><%=nm_marca%></td>	
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center"><b>N/S</b></td>
					<td colspan="4"><%=cd_ns%></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center"><b>Patrim�mio</b></td>
					<td colspan="4"><%=cd_patrimonio%></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center"><b>Motivo</b></td>
					<td colspan="4"><%=motivo%></td>
				</tr>
			<!--Fim da etapa 1.1-->			
			<!--In�cio da etapa 1.2-->				
				<tr style="visibility:<%=visual%>;"><td colspan="100%" align="right"><!--Separador de etapas-->&nbsp;<a href="manut_os_nova.asp?cd_codigo=<%=cd_codigo%>&campo=cd_codigo"> alterar</a></td></tr>
								
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">PROVID�NCIA</td></tr>
				<tr bgcolor="#f5f5f5">
				  	<td  align="center"><b>Data</b></td>
					<td><%=dt_entrada%></td>
					<td  align="center"><b>Avalia��o</b></td>
					
					<td colspan="2"><%=nm_avaliacao%></td>
				</tr>
				<tr bgcolor="#f5f5f5"><td  align="center"><b>Encaminhar para</b></td>
					<td colspan="4"><%=nm_fornecedor%></td>
				</tr>
				<tr bgcolor="#f5f5f5"><td  align="center"><b>Liberado por</b></td>
					<td colspan="4"><%=nm_nome%><%=session("nm_nome")%></td>
				</tr>	
				<tr bgcolor="#f5f5f5">
					<td align="center"><b>Observa��es<br></b></td>
					<td colspan="4"><%=observacoes%></td>
				</tr>
				<tr style="visibility:<%=visual%>;"><td colspan="100%" align="right"><!--Separador de etapas-->&nbsp;<a href="manut_os_nova.asp?cd_codigo=<%=cd_codigo%>&campo=cd_codigo"> alterar</a></td></tr>
				</tr>
				<!--Os novos campos come�am aqui-->
				<%
				sequencia = "2"
				xsql_andamento ="up_andamento_lista @cd_codigo='"&num_os&"',@campo=cd_os, @sequencia='"&sequencia&"'"
					Set rs_andamento = dbconn.execute(xsql_andamento)
					
				If rs_andamento.EOF Then
				'Response.write("fim")
				Else 
				'response.write("ok")
				
				cd_cod = rs_andamento("cd_codigo")
				cod = rs_andamento("cd_cod")
				dt_manut_orc = rs_andamento("dt_manut_orc")
				nm_fornecedor = rs_andamento("nm_fornecedor")
				nm_contato = rs_andamento("nm_contato")
				dt_prev_resposta = rs_andamento("dt_prev_resposta")
				dt_resposta_orc = rs_andamento("dt_resposta_orc")
				nm_orcamento = rs_andamento("nm_orcamento")
				num_os_assist = rs_andamento("num_os_assist")
				dt_prev_retorno = rs_andamento("dt_prev_retorno")
				cd_situacao = rs_andamento("cd_situacao")
				cd_valor_orc = rs_andamento("cd_valor_orc")
					if not cd_valor_orc = "" Then
					cd_valor_orc = formatnumber(cd_valor_orc,2)
					end if
				comentarios = rs_andamento("comentarios")
				dt_retorno = rs_andamento("dt_retorno")
				dt_retorno_un = rs_andamento("dt_retorno_un")
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
				<input type="hidden" name="sequencia" value="<%=int(sequencia)%>">
				<%if esconder = "0" Then
				Else%>
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">MANUTEN��O / OR�AMENTO</td></tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Data manut. / orc.</td>
					<td><input type="text" name="dt_manut_orc" size="11" maxlength="50" class="inputs" value="<%=dt_manut_orc%>"></td> 
					<td  align="center">Assist. T�cnica / Fornecedor</td>
					<input type="hidden" name="cd_fornecedor" value="<%=cd_fornecedor%>">
					<td colspan="2"><%=nm_fornecedor%><%'=cod%></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					
					<td align="center">Contato</td><td><input type="text" name="nm_contato" size="15" class="inputs"  value="<%=nm_contato%>"></td>
					<td  align="center">Previs�o resposta</td>
					<td><input type="text" name="dt_prev_resposta" size="11" maxlength="20" class="inputs" value="<%=dt_prev_resposta%>"></td>
				</tr>
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>
				<tr>
					<td  align="center">Data da resposta do or�amento</td>
					<td colspan="2"><input type="text" name="dt_resposta_orc" size="11" maxlength="20" class="inputs" value="<%=dt_resposta_orc%>"></td>
					<td  align="center">Or�amento</td>
					<%If nm_orcamento = "0" Then%>
					<%ck_mo10 = "selected"
					response.write("zero")%>
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
																<option value="1" <%=ck_mo11%>>Aguardando</option>
																<option value="2" <%=ck_mo12%>>Aprovado</option>
																<option value="0" <%=ck_mo10%>>Reprovado</option>
																<option value="3" <%=ck_mo13%>>Garantia</option>
					</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center">O.S. Assist. T�c.</td>
					<td colspan="2"><input type="text" name="num_os_assist" size="35" maxlength="100" class="inputs" value="<%=num_os_assist%>"></td>
					<td  align="center">Valor</td>
					<td><input type="text" name="cd_valor_orc" size="11" maxlength="20" class="inputs" value="<%=cd_valor_orc%>"></td>
					
				</tr>
				<tr bgcolor="#f5f5f5">
					<td  align="center">Situa��o</td>
					<%strsql ="SELECT * From TBL_situacao Order by nm_situacao"
				  	Set rs_situac = dbconn.execute(strsql)%>
			
					<td colspan="2"><select name="cd_situacao" class="inputs">
									<option value=""></option>
					<%Do while Not rs_situac.EOF%><%strsituacao = rs_situac("cd_codigo")%>
					
					<%if cd_situacao = strsituacao then%><%ck_sit = "selected"%><%end if%>
									<option value="<%=rs_situac("cd_codigo")%>" <%=ck_sit%>><%=rs_situac("nm_situacao")%></option>	
									
					<%ck_sit = ""%>
					<%rs_situac.movenext%>
					<%Loop%>
					</td>
					<td  align="center">Previs�o retorno/entrega</td>
					<td><input type="text" name="dt_prev_retorno" size="11" maxlength="20" class="inputs" value="<%=dt_prev_retorno%>"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Laudo de <br>manuten��o</td>
					<td colspan="4">
						<textarea cols="40" rows="2" name="comentarios" class="inputs"><%=comentarios%></textarea></td>
				</tr>
				<tr>
					<td  align="center">Data de entrega/retorno</td>
					<td colspan="2"><input type="text" name="dt_retorno" size="11" maxlength="20" class="inputs" value="<%=dt_retorno%>"></td>
				</tr>
				<tr>
					<td  align="center">Data de envio para Hospital</td>
					<td colspan="2"><input type="text" name="dt_retorno_un" size="11" maxlength="20" class="inputs" value="<%=dt_retorno_un%>"></td>
				</tr>
				
				
				<%end if%>
				<tr style="visibility:<%=visual%>;">
					<td colspan="100%" align="center">
						<input type="hidden" name="cd_codigo" value="<%=cd_cod%>">
						<input type="hidden" name="acao" value="<%=acao%>">
						<input type="submit" value="Enviar"><%'=acao%>						
					</td>
				</tr>
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



</BODY>
</HTML>



