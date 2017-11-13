<!--include file="../includes/util.asp"-->
<!--include file="../includes/inc_open_connection.asp"-->
<!--include file="../includes/estilos_css.htm"-->

<%cd_user =  session("cd_codigo")
pasta_loc = "Manutencao2\"
arquivo_loc = "manutencao_nova_3.asp"

aviso = request("aviso")
jan = request("jan")
os_alt = request("os_alt")
				
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
<script language="Javascript" TYPE="text/javascript" >
	/*** Arquivos componentes da ordem de serviço:
		* ajax_os_natureza
		* ajax_os_patrimonio_busca
		* ajax_os_patrimonio_compra
		* ajax_os_pat_NS_mostra
	*/
	
	function seleciona_natureza(cod){
		//*** Seleciona a natureza da ordem de serviço ***
		loadXMLDoc("manutencao_2/ajax2/ajax_os_natureza.asp?natureza_os="+cod,function(){				
			if (xmlhttp.readyState==4 && xmlhttp.status==200){ document.getElementById("divFase0").innerHTML = xmlhttp.responseText;}
		  }
			);
			// limpa campos do equipamento/ otica/ instrumento
			document.getElementById("cd_patrimonio").value='';
			document.getElementById("cd_equipamento").value='';
			document.getElementById("cd_marca").value='';
			document.getElementById("cd_especialidade").value='';
			document.getElementById("cd_ns").value='';
			document.getElementById("num_qtd").value='';
		}	
	
	function seleciona_tipo_objeto_m(cod){
		//*** Seleciona entre equip, ótica, instr e material ***
		loadXMLDoc("manutencao_2/ajax2/ajax_os_patrimonio_busca.asp?tipo_item="+cod,function(){				
			if (xmlhttp.readyState==4 && xmlhttp.status==200){ document.getElementById("divFase1").innerHTML = xmlhttp.responseText;}
		  }
			);
			// limpa campos do equipamento/ otica/ instrumento
			document.getElementById("cd_patrimonio").value='';
			document.getElementById("cd_equipamento").value='';
			document.getElementById("cd_marca").value='';
			document.getElementById("cd_especialidade").value='';
			document.getElementById("cd_ns").value='';
			document.getElementById("num_qtd").value='';
		}
	
	function seleciona_tipo_objeto_c(cod){
		//*** Seleciona entre equip, ótica, instr e material ***
		loadXMLDoc("manutencao_2/ajax2/ajax_os_patrimonio_compra.asp?tipo_item="+cod,function(){				
			//alert(xmlhttp.readyState+'-'+xmlhttp.status);
			if (xmlhttp.readyState==4 && xmlhttp.status==200){ document.getElementById("divFase1").innerHTML = xmlhttp.responseText;}
		  }
			);
			// limpa campos do equipamento/ otica/ instrumento
			document.getElementById("cd_patrimonio").value='';
			document.getElementById("cd_equipamento").value='';
			document.getElementById("cd_marca").value='';
			document.getElementById("cd_especialidade").value='';
			document.getElementById("cd_ns").value='';
			document.getElementById("num_qtd").value='';
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
	
	function mostra_NS(tipo,codigo_ns){
		//*** Seleciona o numero do patrimonio ou o numero de serie ***
		loadXMLDoc("manutencao_2/ajax2/ajax_os_pat_NS_mostra.asp?tipo_pat="+tipo+"&no_patrimonio=&no_serie="+codigo_ns,function()
			{				
			
			if (xmlhttp.readyState==4 && xmlhttp.status==200 || xmlhttp.readyState==4 && xmlhttp.status==500)
		    		{
		    			document.getElementById("divPatNS_mostra").innerHTML = xmlhttp.responseText;
		    		}
		  		}
			);
		// limpa campos do equipamento/ otica/ instrumento
			document.getElementById("numPAT").value='';
			document.getElementById("cd_patrimonio").value='';
			document.getElementById("cd_equipamento").value='';
			document.getElementById("cd_marca").value='';
			document.getElementById("cd_especialidade").value='';
			document.getElementById("cd_ns").value='';
			document.getElementById("num_qtd").value='';
	}
	
	function alimenta_cd_pat(cd_pat){
	document.getElementById("cd_patrimonio").value=cd_pat;
	}	
	function alimenta_cd_equipamento(cd_equipamento){
	document.getElementById("cd_equipamento").value=cd_equipamento;
	}
	function alimenta_cd_marca(cd_marca){
	document.getElementById("cd_marca").value=cd_marca;
	}
	function alimenta_cd_especialidade(cd_especialidade){
	document.getElementById("cd_especialidade").value=cd_especialidade;
	}
	function alimenta_cd_qtd(num_qtd){
	document.getElementById("num_qtd").value=num_qtd;
	}
	
	function alimenta_info_pat(cd_pat,no_patrimonio,cd_equipamento,cd_marca,cd_especialidade,cd_ns,num_qtd){
	document.getElementById("cd_patrimonio").value=cd_pat;
	document.getElementById("cd_equipamento").value=cd_equipamento;
	document.getElementById("cd_marca").value=cd_marca;
	document.getElementById("cd_especialidade").value=cd_especialidade;
	document.getElementById("cd_ns").value=cd_ns;
	document.getElementById("num_qtd").value=num_qtd;
	}
	
</script>
<!--include file="../js/ajax.js"-->

<!--#include file="../js/ajax2.js"--> 
<%esconder = "1"%>



<!--#include file="../includes/arquivo_loc.asp"-->

<br>
<!--table align="center" border="1" cellpadding="1" cellspacing="1">
	<tr>
		<td-->
			<table width="600" border="1" align="center" style="border-collapse:collapse;"  classeeeee="textopequeno">
				<tr style="background-color:#b7d7ee;">		
					<td class="txt_cinza" colspan="6">
					<b>Manutenção &raquo; - <font color="red">Listagem</font></b></td>
				</tr>
				<%if jan <> 1 then '*** Caso esta pagina seja aberta em uma janela, não mostra este bloco ***%>
				<tr><td colspan="4"><img src="imagens/blackdot.gif" alt="" width="600" height="5" border="0"></td></tr>
					<form action="manutencao2.asp" name="busca" id="busca">
					<input type="hidden" name="tipo" value="nova2">
				<tr><td colspan="4" style="background-color:#e6fae2;">NOVA O.S.: 
					<input type="text" name="cd_unidade2" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" class="inputs">
					<input type="text" name="num_os2" size="10" maxlength="10" onBlur="submit();" class="inputs">
					<input type="submit" value="ok"></td></tr>
					</form>
				<%end if%>
				<tr><td colspan="4"><img src="imagens/blackdot.gif" alt="" width="600" height="5" border="0"></td></tr>
				<tr><td colspan="4">&nbsp;</td></tr>
				<%if session_cd_usuario <> usuario_restrito then%>
				<tr>
				    <td colspan="4" class="textopequeno" bgcolor="#f5f5f5">
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
if cd_unidade <> "" AND num_os <> "" OR os_alt = 1 Then


				
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
						<input type="hidden" name="tipo" value="nova3">
						<input type="hidden" name="sequencia" value="1">
						<input type="hidden" name="jan" value="<%=jan%>">
						<input type="hidden" name="acao" value="<%=acao%>">
						<input type="hidden" name="filtro" value="<%=filtro%>">
						<input type="hidden" name="cd_codigo" value="<%=cd_codigo%>">						
						<input type="hidden" name="cd_liberacao" value="<%=session("nm_nome")%>">
						
				<tr bgcolor="#b4b4b4">
					<td colspan="4" align="center" class="textopadrao">SOLICITAÇÃO</td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center" valign="bottom">O.S.<br><img src="imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td align="center" valign="bottom">
						<%if cd_os = "" Then%>
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
						<br><img src="imagens/px.gif" alt="" width="120" height="1" border="0">
					</td>
					<td align="center" valign="bottom"><b>NATUREZA:</b><br><img src="imagens/px.gif" alt="" width="80" height="1" border="0"></td>
					<td align="center" valign="bottom">
						<!-- Chama a página manutencao_2/ajax/ajax_os_natureza.asp -->
						<input type="radio" name="cd_natureza_os" value="1" onMouseup="seleciona_natureza('m');" <%if int(cd_natureza_os) = 1 then%>checked<%end if%>>Manutenção &nbsp; &nbsp; &nbsp; 
						<input type="radio" name="cd_natureza_os" value="2" onMouseup="seleciona_natureza('c');" <%if int(cd_natureza_os) = 2 then%>checked<%end if%>>Compra &nbsp; &nbsp; &nbsp; 
						<input type="radio" name="cd_natureza_os" value="3" onMouseup="seleciona_natureza('p');" <%if int(cd_natureza_os) = 3 then%>checked<%end if%>>Preventiva
							<%if cd_os <> "" AND cd_user = "10" OR cd_os <> "" AND cd_user = "46" Then%>
							<img src="../../imagens/ic_del.gif" alt="" width="13" height="14" border="0" onClick="javascript:JsOSDelete('<%=cd_os%>')"><%end if%>
							<br><img src="imagens/px.gif" alt="" width="300" height="1" border="0"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Data</td>
					<td colspan="3"><input type="text" name="dt_os_dia" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=dt_os_dia%>" class="inputs">/
									<input type="text" name="dt_os_mes" size="2" maxlength="2"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 2)" onFocus="PararTAB(this);" value="<%=dt_os_mes%>" class="inputs">/
									<input type="text" name="dt_os_ano" size="4" maxlength="4"  onKeyPress="ChecarTAB();" onKeyUp="Mostra(this, 4)" onFocus="PararTAB(this);" value="<%=dt_os_ano%>" class="inputs">
									<!--input type="text" name="dt_os" size="11" maxlength="50" class="inputs" value="<%=dt_os%>" onFocus="nextfield ='nm_solicitante';"--></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Solicitante</td>
					<td colspan="3"><input type="text" name="nm_solicitante" size="40" maxlength="40" class="inputs" value="<%=nm_solicitante%>" onFocus="nextfield ='cd_unidade_os';"></td>
				</tr>
				<tr bgcolor="#f5f5f5">
					<td align="center">Unidade</td>
					<td colspan="3">
						<%strsql ="SELECT * FROM TBL_unidades WHERE cd_codigo = '"&cd_unidade&"' AND cd_status = 1 AND cd_hospital >= 1 order by cd_codigo"
						Set rs_uni = dbconn.execute(strsql)%>
						<%while not rs_uni.EOF
							cd_uni = int(rs_uni("cd_codigo"))
							nm_sigla = rs_uni("nm_sigla")
							nome_unidade = rs_uni("nm_unidade")%>
							
							<b><%=nome_unidade%></b>
						<%rs_uni.movenext
						wend%>							
					</td>					
				</tr>
				<tr><td colspan="4">&nbsp;</td></tr>
				<tr bgcolor="#b4b4b4">
					<td colspan="4" align="center" class="textopadrao"><b>PATRIMÔNIO / INSTRUMENTO / MATERIAL</b></td>
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
						</table>
					</td>
				</tr>
				<%end if%>
				<tr bgcolor="#f5f5f5">
					<td>&nbsp;</td>
					<td colspan="3">
						<span id="divFase0">
							<span>
								<%if cd_os = "" Then%><p align="center" style="color:red;font-weight:bold;">*** Selecione primeiro a natureza da ordem de serviço. ***
								<%else%><p align="center" style="color:black;font-weight:bold;"*** >Para alterar o item acima, clique novamente na natureza da O.S. ***
								<%end if%> 
							</span>
								&nbsp;
							<!--span id="divFase1"-->
									
							<!--/span-->
						</span>
						
						
					</td>
				</tr>
				
				<tr bgcolor="#b4b4b4">
					<td colspan="4" align="center" class="textopadrao">&nbsp;
					<input type="hidden" name="cd_patrimonio" value="<%=cd_patrimonio%>" id="cd_patrimonio">
					<input type="hidden" name="cd_equipamento" value="<%=cd_equipamento%>" id="cd_equipamento">
					<input type="hidden" name="cd_marca" value="<%=cd_marca%>" id="cd_marca">
					<input type="hidden" name="cd_especialidade" value="<%=cd_especialidade%>" id="cd_especialidade">
					<input type="hidden" name="cd_ns" value="<%=cd_ns%>" id="cd_ns">
					<input type="hidden" name="num_qtd" value="<%=num_qtd%>" id="num_qtd">
					<!--input type="text" name="array_manut" id="array_manut" value="" size="50" onclick="teste();"--></td>
				</tr>				
				<tr bgcolor="#f5f5f5">
					<td align="center">Motivo</td>
					<td colspan="3">
						<textarea cols="65" rows="2" name="motivo" class="inputs" onFocus="nextfield ='dt_entrada';"><%=motivo%></textarea>
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
					<td colspan="1"><img src="../imagens/px.gif" alt="" width="5" height="1" border="0"></td>
					<!--td colspan="2"><img src="../imagens/blackdot.gif" alt="" width="300" height="1" border="0"></td-->
				</tr>			
			</table>

		<!--/td>
	</tr>
</table-->
<%end if%>
	<SCRIPT LANGUAGE="javascript">
		function JsOSDelete(cod){
			if (confirm ("Tem certeza que deseja excluir a Ordem de Serviço: <%=zero(cd_unidade)%>.<%=num_os%>?")){
				document.location.href('<%=caminho%>acoes/manutencao_acao.asp?cd_codigo_os='+cod+'&acao=3&filtro=apaga');
			}
		}
	</SCRIPT>






