<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<!--#include file="../includes/estilos_css.htm"-->
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

<script language="JavaScript" type="text/javascript" src="../js/richtext.js"></script>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../js/formValidator.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript" SRC="../js/mascarainput.js"></SCRIPT>
<script language="javascript">
<!--
  /*function mOvr(src,clrOver) {
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
  }*/

// -->	
</script>
<script language="JavaScript">
function validar_os(shipinfo){
	if (shipinfo.dt_os_dia.value.length==""){window.alert ("O dia de abertura da O.S. não foi preenchido.");shipinfo.dt_os_dia.focus();return (false);
	}
	if (shipinfo.dt_os_mes.value.length==""){window.alert ("O mês de abertura da O.S. não foi preenchido.");shipinfo.dt_os_mes.focus();return (false);
	} 
	if (shipinfo.dt_os_ano.value.length==""){window.alert ("O ano  de abertura da O.S. não foi preenchido.");shipinfo.dt_os_ano.focus();return (false);
	} 
	if (shipinfo.nm_solicitante.value.length==""){window.alert ("O solicitante não foi informado.");shipinfo.nm_solicitante.focus();return (false);
	}
	if (shipinfo.cd_unidade_os.value.length==""){window.alert ("Selecione uma unidade.");shipinfo.cd_unidade_os.focus();return (false);
	}
	if (shipinfo.cd_equipamento.value.length==""){window.alert ("Informe:  Equipamento, Ótica, Instrumento ou Material");shipinfo.dt_os_dia.focus();return (false);
	}
	if (shipinfo.num_qtd.value.length==""){window.alert ("informe a quantidade.");shipinfo.dt_os_dia.focus();return (false);
	}
	if (shipinfo.motivo.value.length==""){window.alert ("Qual motivo da solicitação?");shipinfo.motivo.focus();return (false);
	}
	return (true);}
</script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
nextfield = "dt_os_dia"; // nome do primeiro campo do site
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

<%cd_user =  session("cd_codigo")
aviso = request("aviso")
jan = request("jan")
				
cd_codigo = request("cd_codigo")
cd_unidade = request("cd_unidade2")
	if cd_unidade = "" Then
		cd_unidade = "00"
	end if
num_os = request("num_os2")
campo = request("campo")
ordem = request("ordem")
list = request("list")
lista = "1"

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
acao = "1"
filtro = "cadastra_os"

if jan = 1 then
	caminho = ""
else
	caminho = "manutencao_2/"
	jan = 0
end if
%>


<!--#include file="../js/ajax.js"-->
<!--#include file="js/manutencao.js"-->
<%esconder = "1"%>




<br>
<!--table align="center" border="1" cellpadding="1" cellspacing="1">
	<tr>
		<td-->
			<table width="600" border="1" cellspacing="1" cellpadding="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#efefef" class="textopequeno">
			<%if session("cd_codigo") = 46 then%>
				<tr><td colspan="2" align="center"><span style="font-size=8px;color:red;">manutencao_nova.asp</span></td></tr>
			<%end if%>
				<tr>		
					<td class="txt_cinza" colspan="6">
					<b>Manutenção &raquo; - <font color="red">Listagem</font></b><br><br></td>
					<!--td class="txt_cinza" align="right" valign="top">Busca O.S.: &nbsp;</td>
					<td class="txt_cinza" colspan="2">
					<form action="ordemservico_andamento.asp" name="busca" id="busca">
					<input type="text" name="cd_codigo" size="6" maxlength="6">
					<input type="hidden" name="campo" value="num_os">
					<input type="hidden" name="ordem" value="1">		
					<input type="submit" value="ok">
					</form>
					</td-->
				</tr>
				<%if jan <> 1 then%>
				<tr><td colspan="2"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
					<form action="manutencao2.asp" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova2">
				<tr><td colspan="2">NOVA O.S.: <input type="text" name="cd_unidade2" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);"><input type="text" name="num_os2" size="10" maxlength="10" onBlur="submit();"><input type="submit" value="ok"></td></tr>
					</form>
				<%end if%>
				<tr><td colspan="2"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
				<%if session_cd_usuario <> usuario_restrito then%>
				<tr>
				    <td colspan="2" class="textopequeno" bgcolor="#f5f5f5">
						<!--/ <a href="manut_os_nova.asp">Nova O.S. </a-->
					    / <a href="manutencao2.asp?tipo=pendencias&filtro=pendentes">Pendentes</a> 
						/ <a href="manutencao2.asp?tipo=pendencias&filtro=encerradas&strtop=20">Encerradas</a>
						/ <a href="manutencao2.asp?tipo=pendencias&filtro=todas&strtop=20">Todas</a>
						/ <a href="manutencao2.asp?tipo=relatorio">Buscas</a> /
					</td>
				</tr>					
				<%end if%>
				
<%'***************************************
'*** Verifica a existência da Unidade. ***
'*****************************************
if num_os <> "" Then
	xsql = "SELECT * FROM TBL_Unidades where cd_codigo="&cd_unidade&" and cd_status = 1"
	Set rs_unidade = dbconn.execute(xsql)
		do while not rs_unidade.EOF	
			cd_unidade_verifica = rs_unidade("cd_codigo")
		rs_unidade.movenext
		loop
			if cd_unidade_verifica = ""  then 'Avisa que a unidade não é válida
				response.redirect("manutencao2.asp?tipo=nova2&aviso=Unidade inválida!")
				
			End if
end if

'*************************************
'*** Verifica a existência da O.S. ***
'*************************************
if num_os <> "" Then
	xsql = "SELECT * FROM TBL_OrdemServico2 where num_os="&num_os&" and cd_unidade="&cd_unidade&""
	Set rs_verifica = dbconn.execute(xsql)
		do while not rs_verifica.EOF
		num_os_verifica = rs_verifica("num_os")
		cd_codigo_verifica = rs_verifica("cd_codigo")
		rs_verifica.movenext
		loop
			if num_os_verifica = int(num_os) then 'Redireciona à O.S. existente
				'response.redirect("manut_os_andamento.asp?cd_codigo="&cd_codigo_verifica&"&campo=cd_codigo")
				response.redirect("manutencao2.asp?tipo=andamento&cd_codigo="&cd_codigo_verifica&"&campo=cd_codigo")
			End if
end if
'*************************************



				
				if cd_codigo <> "" Then
				'xsql ="up_os_lista @cd_unidade='"&cd_unidade&"',@cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
				xsql ="up_os_lista2 @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
				Set rs = dbconn.execute(xsql)
				
				cd_os = rs("cd_codigo")
				num_os = rs("num_os")
				cd_natureza_os = rs("cd_natureza_os")
					if cd_natureza_os = "" then
						 cd_natureza_os = "0"
					end if
				
				dt_os = rs("dt_os")
					dt_os_dia = day(dt_os)
					dt_os_mes = month(dt_os)
					dt_os_ano = year(dt_os)
					
				'nm_unidade = rs("nm_unidade"
				cd_unidade = rs("cd_unidade")
				cd_unidade_os = rs("cd_unidade_os")
				
				cd_especialidade = rs("cd_especialidade")
				nm_especialidade = rs("nm_especialidade")
				
				cd_equipamento = rs("cd_equipamento")
				nm_equipamento = rs("nm_equipamento")
				nm_equipamento_novo = rs("nm_equipamento_novo")
				cd_ns = rs("cd_ns")
				cd_patrimonio = rs("cd_patrimonio")
				no_patrimonio = rs("no_patrimonio")
				no_pat_old = rs("no_pat_old")
				cd_marca = rs("cd_marca")
				nm_marca = rs("nm_marca")
				num_qtd = rs("num_qtd")
				
				dt_entrada = rs("dt_entrada")
				cd_avaliacao = rs("cd_avaliacao")
				'nm_avaliacao = rs("nm_avaliacao")
				nm_fornecedor = rs("nm_fornecedor")
				cd_fornecedor = rs("cd_fornecedor")
				cd_liberacao = rs("cd_liberacao")
			
				cd_codigo = rs("cd_codigo")
			
				
				
				motivo = rs("motivo")
				nm_solicitante = rs("nm_solicitante")
				observacoes = rs("observacoes")
				
				acao = "2"
				filtro = "modifica_os"
				sequencia = "1"
				fecha = rs("fecha")
				End if%>
			<!--Início da etapa 1.1-->
							
					<form method="post" action="<%=caminho%>acoes/manutencao_acao.asp" name="form" id="form" onsubmit="return validar_os(document.form);" >
						<input type="hidden" name="tipo" value="nova">
						<input type="hidden" name="sequencia" value="1">
						<input type="hidden" name="jan" value="<%=jan%>">
						<input type="hidden" name="acao" value="<%=acao%>">
						<input type="hidden" name="filtro" value="<%=filtro%>">
						<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">						
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
						
				<tr bgcolor="#b4b4b4">
					<td  colspan="2" align="center" class="textopadrao">SOLICITAÇÃO</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">O.S.</td>
					<td colspan="1"><%if cd_os = "" Then%>
										<span style="font-size:13px; font-weight:bold;">
											<%if aviso <> "" Then%>
												<b style=" color:red;"><%=aviso%></b>
											<%else%>
												<%=cd_unidade%>.<%=num_os%>
											<%end if%>
										</span>
										<input type="hidden" name="cd_unidade" value="<%=cd_unidade%>">
										<input type="hidden" name="num_os" value="<%=num_os%>">
									<%else%>
										<input type="text" name="cd_unidade" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=cd_unidade%>"><input type="text" name="num_os" size="5" maxlength="4" class="inputs" value="<%=num_os%>" onFocus="nextfield ='dt_os';">&nbsp;
									<%end if%>
									&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;	&nbsp;	
									<!--input type="text" name="digito" size="1" maxlength="3" class="inputs" value=""></td>
									<td colspan="3"--> 
									NATUREZA: <input type="radio" name="cd_natureza_os" value="1" onMouseup="mostra_os_natureza('m');" <%if int(cd_natureza_os) = 1 then%>checked<%end if%>>Manutenção &nbsp; &nbsp; &nbsp; 
											  <input type="radio" name="cd_natureza_os" value="2" onMouseup="mostra_os_natureza('c');" <%if int(cd_natureza_os) = 2 then%>checked<%end if%>>Compra &nbsp; &nbsp; &nbsp; 
											  <input type="radio" name="cd_natureza_os" value="3" onMouseup="mostra_os_natureza('p');" <%if int(cd_natureza_os) = 3 then%>checked<%end if%>>Preventiva &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
											  <%if cd_os <> "" AND cd_user = "10" Then%><img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsOSDelete('<%=cd_os%>')"><%end if%></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Data</td>
					<td colspan="1"><input type="text" name="dt_os_dia" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=dt_os_dia%>" class="inputs">/
									<input type="text" name="dt_os_mes" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=dt_os_mes%>" class="inputs">/
									<input type="text" name="dt_os_ano" size="4" maxlength="4"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=dt_os_ano%>" class="inputs">
									<!--input type="text" name="dt_os" size="11" maxlength="50" class="inputs" value="<%=dt_os%>" onFocus="nextfield ='nm_solicitante';"--></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Solicitante</td>
					<td colspan="1"><input type="text" name="nm_solicitante" size="40" maxlength="40" class="inputs" value="<%=nm_solicitante%>" onFocus="nextfield ='cd_unidade_os';"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Unidade</td>
						<%strsql ="SELECT * FROM TBL_unidades WHERE cd_status = 1 AND cd_hospital >= 1 order by cd_codigo"
						Set rs_uni = dbconn.execute(strsql)%>
					<td valign="top" colspan="1">
					<span id='divUnid'>
					<select name="cd_unidade_os" class="inputs">
					<option value="">Selecione</option>
						<%while not rs_uni.EOF
						cd_uni = rs_uni("cd_codigo")
						nm_sigla = rs_uni("nm_sigla")
						nome_unidade = rs_uni("nm_unidade")
						%>						
					<option value="<%=cd_uni%>" <%if int(cd_unidade_os) = cd_uni then%>selected<%end if%>><%'=nm_sigla%><%=nome_unidade%></option>
						<%rs_uni.movenext
						ck_uni = ""
						'salto = salto + 1
						wend
						%>
					</select>
					</span>
					<!--img src="../imagens/reload6.gif" alt="Atualizar" border="0" onClick="fill_unid();"--> 
					<!--a href="javascript: void(0);" onclick="window.open('janelas_cadastros/unidade_cad.asp?janela=pop_up&metodo=fecha', 'principal', 'width=445, height=185'); return false;"><img src="../imagens/ic_novo.gif" alt="inserir" border="0"></a-->					
					</td>
					<td colspan="1"bgcolor="#f5f5f5" valign="bottom" align="center"><!--a href="#" onclick="window.open('adm/equipamento_cad.asp', 'Janela','width=430,height=500')">+<p></p--></td>
				</tr>
				<tr><td colspan="1">&nbsp;</td></tr>
				<tr bgcolor="#b4b4b4">
					<td  colspan="2" align="center" class="textopadrao">PATRIMÔNIO / INSTRUMENTO / MATERIAL</td>
				</tr>
				<%if cd_os <> "" Then%>
				<tr bgcolor="#f5f5f5">
					<td align="center">Item atual</td>
					<td>
						<table>
							<tr>
								<td><b>Pat:</b></td>
								<td style="color:#339900;"><b><%=no_patrimonio%></b></td>
								<td><b>Qtd:</b></td>
								<td><%=num_qtd%></td>
								<!--td colspan="2"><input type="text" name="cd_patrimonio" value="<%=cd_patrimonio%>"></td-->
							</tr>
						  	<tr>
								<td style="width:45px;"><b>Item:</b></td>
								<td style="width:150px;"><%=nm_equipamento_novo%></td>
								<td style="width:45px;"><b>Especialidade:</b></td>
								<td style="width:150px;"><%=nm_especialidade%></td>			
							</tr>
							<tr style="color:gray;">
								<td style="width:45px;"><b>Antigo:</b></td>
								<td style="width:150px;"><%=nm_equipamento%></td>
								<td style="width:45px;"><b>Patrimonio:</b></td>
								<td style="width:150px;"><%=no_pat_old%></td>			
							</tr>
							<tr>
								<td style="width:45px;"><b>Marca:</b></td>
								<td style="width:150px;"><%=nm_marca%></td>
								<td style="width:45px;"><b>N/S:</b></td>
								<td style="width:150px;"><%=cd_ns%></td>
							</tr>
								<!--input type="hidden" name="cd_patrimonio_atual" value="<%=cd_patrimonio%>">
								<input type="hidden" name="cd_equipamento_atual" value="<%=cd_equipamento%>">
								<input type="hidden" name="cd_marca_atual" value="<%=cd_marca%>">
								<input type="hidden" name="cd_especialidade_atual" value="<%=cd_especialidade%>">
								<input type="hidden" name="cd_ns_atual" value="<%=cd_ns%>">
								<input type="hidden" name="num_qtd_atual" value="<%=num_qtd%>"-->
						</table>
					</td>
				</tr>
				<%end if%>
				<tr bgcolor="#f5f5f5">
					<td>&nbsp;</td>
					<td colspan="1"><span id="divFase0">				
						<span>
						<%if cd_os = "" Then%><p align="center" style="color:red;font-weight:bold;">Selecione primeiro a natureza da ordem de serviço. &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;
						<%else%><p align="center" style="color:black;font-weight:bold;">Para alterar o item acima, clique novamente na natureza da O.S. &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;
						<%end if%> </span>
						<!--<input type="radio" name="cd_tipo_item" value="e" onFocus="mostra_itens('e');">Equipamento &nbsp; &nbsp; &nbsp; &nbsp; 
						<input type="radio" name="cd_tipo_item" value="o" onFocus="mostra_itens('o');">Ótica &nbsp; &nbsp; &nbsp; &nbsp; 
						<input type="radio" name="cd_tipo_item" value="i" onFocus="mostra_itens('i');">Instrumento &nbsp; &nbsp; &nbsp; &nbsp; 
						<input type="radio" name="cd_tipo_item" value="m" onFocus="mostra_itens('m');">Material
						-->
							<br>&nbsp;
							<span id="divFase1">
								<input type="hidden" name="cd_patrimonio" value="<%=cd_patrimonio%>">
								<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>">
								<input type="hidden" name="cd_marca" value="<%=cd_marca%>">
								<input type="hidden" name="cd_especialidade" value="<%=cd_especialidade%>">
								<input type="hidden" name="cd_ns" value="<%=cd_ns%>">
								<input type="hidden" name="num_qtd" value="<%=num_qtd%>">
							</span>
						</span>
					</td>
				</tr>
				
				<tr bgcolor="#b4b4b4">
					<td colspan="2" align="center" class="textopadrao">&nbsp;</td>
				</tr>				
				<tr bgcolor="#f5f5f5">
					<td align="center">Motivo</td>
					<td colspan="1">
						<textarea cols="70" rows="2" name="motivo" class="inputs" onFocus="nextfield ='dt_entrada';"><%=motivo%></textarea>
					</td>
				</tr>
				
				<tr>
					<td colspan="2" align="center">
					<br><%If acao = "1" Then%><input  name="ok" type="submit" value="Cadastrar O.S."><%Else%>
					<input type="submit" value="Modificar O.S."><%if jan <> 1 then%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="manutencao2.asp?tipo=andamento&cd_codigo=<%=cd_codigo%>&campo=<%=campo%>">Voltar</a><%end if%><%End if%>
																										
					</td>
				</tr>
				<tr>
					<td><img src="../imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td colspan="1"><img src="../imagens/px.gif" alt="" width="580" height="1" border="0"></td>
					<!--td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td-->
				</tr>			
			</table>

		<!--/td>
	</tr>
</table-->

	<SCRIPT LANGUAGE="javascript">
		function JsOSDelete(cod){
			if (confirm ("Tem certeza que deseja excluir a Ordem de Serviço: <%=zero(cd_unidade)%>.<%=num_os%>?")){
				document.location.href('<%=caminho%>acoes/manutencao_acao.asp?cd_codigo_os='+cod+'&acao=3&filtro=apaga');
			}
		}
	</SCRIPT>






