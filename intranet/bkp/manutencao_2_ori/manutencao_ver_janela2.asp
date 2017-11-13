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
<script language="JavaScript">
	function mouseFora(){
		window.close();
	}
</script>
<%esconder = "1"%>
<BODY onfocusout="mouseFora();">
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

xsql ="up_os_lista2 @cd_codigo='"&cd_codigo&"',@campo='"&campo&"', @ordem='"&ordem&"', @top='"&strtop&"',@compare='"&compare&"', @outros='"&stroutros&"'"
Set rs = dbconn.execute(xsql)
If rs.EOF Then%>
			
				<%'end if
			registros = 0
			Else
			Do while not rs.EOF
				cd_cod_andamento = rs("cd_codigo")
				cd_unidade = rs("cd_unidade")
				num_os = rs("num_os")
				dt_os = rs("dt_os")
				'nm_unidade = rs("nm_unidade")
				cd_unidade = rs("cd_unidade")
				nm_unidade = rs("nm_sigla")
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
				no_patrimonio = rs("no_patrimonio")
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
					<b>Manutenção &raquo; - <span style="color:black;">Visualiza ordem de serviço</span></b></td>
				</tr>
				<tr id="ok_print"><td colspan="100%"><img src="imagens/blackdot.gif" alt="" width="100%" height="1" border="0"></td></tr>										
			<!--Início da etapa 1.1-->			
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
					<td colspan="4"><%=no_patrimonio%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Motivo</b></td>
					<td colspan="4"><%=motivo%></td>
				</tr>
			<!--Fim da etapa 1.1-->			
			<!--Início da etapa 1.2-->				
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">PROVIDÊNCIA</td></tr>
				<tr bgcolor="#<%=bg_cor%>">
				  	<td  align="center"><b>Data envio</b></td>
					<td><%=dt_entrada%></td>
					<td  align="center"><b>Avaliação</b></td>
					<td colspan="2">
						<%strsql ="SELECT * From TBL_avaliacao where cd_codigo="&cd_avaliacao&""
					  	Set rs_aval = dbconn.execute(strsql)%>						
						<%Do while Not rs_aval.EOF%>
							<%=rs_aval("nm_avaliacao")%>
						<%rs_aval.movenext
						Loop%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Encaminhar para</b></td>						
						<%strsql ="TBL_Fornecedor order by nm_fornecedor"
					  	Set rs_forn = dbconn.execute(strsql)%>
					<td colspan="4"><%=nm_fornecedor%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>"><td  align="center"><b>Cadastrado por</b></td>
					<td colspan="4"><%=cd_liberacao%><%'=nm_nome%><%'=session("nm_nome")%></td>
				</tr>	
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Ocorrência</b><br></td>
					<td colspan="4"><%=observacoes%></td>
				</tr>				
				<tr><td colspan="5">&nbsp;</td></tr>
				<!--Os novos campos começam aqui-->
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
				nm_contato = rs_andamento("nm_contato")
				dt_prev_resposta = rs_andamento("dt_prev_resposta")
				dt_resposta_orc = rs_andamento("dt_resposta_orc")
				nm_orcamento = rs_andamento("nm_orcamento")
				num_os_assist = rs_andamento("num_os_assist")
				dt_prev_retorno = rs_andamento("dt_prev_retorno")
				cd_situacao = rs_andamento("cd_situacao")
				cd_gestao = rs_andamento("cd_gestao")
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
				<%if session_cd_usuario <> usuario_restrito then%>
				<%if esconder = "0" Then
				Else%>
				<tr bgcolor="#b4b4b4">
					<td colspan="100%" align="center" class="textopadrao">MANUTENÇÃO / ORÇAMENTO <b style="color:#b4b4b4;"><%=cd_cod_andamento%></b></td>
				</tr>					
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Data manut./orc.</b></td>
					<td><%=dt_manut_orc%></td> 
					<td  align="center"><b>Assist. Técnica / Fornecedor</b></td>
					<!--input type="hidden" name="cd_fornecedor" value="<%=cd_fornecedor%>"-->
					<td colspan="2"><%=nm_fornecedor%><%'=cod%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					
					<td align="center"><b>Contato</b></td>
					<td><%=nm_contato%></td>
					<td  align="center"><b>Previsão resposta</b></td>
					<td><%=dt_prev_resposta%></td>
					<td>&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">					
					<td align="center"><b>Motivo</b></td>
					<td  colspan="2">
					<%if nm_natureza_defeito = 1 then %>
						Motivo desconhecido
					<%elseif nm_natureza_defeito = 2 then%>
						Falha do operador
					<%elseif nm_natureza_defeito = 3 then%>
						Desgaste por uso
					<%elseif nm_natureza_defeito = 4 then%>
						Quebra por terceiros
					<%elseif nm_natureza_defeito = 5 then%>
						Solicitação de compras por terceiros
					<%end if%>
					</td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr><td colspan="100%"><!--Separador de etapas-->&nbsp;</td></tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Data da resposta do orçamento</b></td>
					<td colspan="2"><%=dt_resposta_orc%></td>
					<td  align="center"><b>Orçamento</b></td>
					<td><%if nm_orcamento = "0" Then%>
						Reprovado
					<%elseif nm_orcamento = "1" Then%>
						Aguardando
					<%elseif nm_orcamento = "2" Then%>
						Aprovado
					<%elseif nm_orcamento = "3" Then%>
						Garantia
					<%End if%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>O.S. Assist. Téc.</b></td>
					<td colspan="2"><%=num_os_assist%></td>
					<td colspan="2">&nbsp;</td>
					
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center"><b>Unidade</b></td>
					<td colspan="2"><%=cd_qtd_valor%> &nbsp; &nbsp; &nbsp; <b>Valor unitário R$</b> <%=cd_valor_unit%></td>
					<td  align="center"><b>Valor Total R$</b></td>
					<td><%=cd_valor_orc%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Situação</b></td>
						<%strsql ="SELECT * From TBL_situacao where cd_codigo="&cd_situacao&" Order by nm_situacao"
					  	Set rs_situac = dbconn.execute(strsql)%>
						<td colspan="2">
						<%Do while Not rs_situac.EOF%>
							<%=rs_situac("nm_situacao")%>
						<%rs_situac.movenext
						Loop%></td>
					<td  align="center"><b>Previsão retorno/entrega</b></td>
					<td><%=dt_prev_retorno%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td align="center" valign="top"><b>Laudo de <br>manutenção</b></td>
					<td colspan="4"><%=comentarios%></td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Gestão</b></td>
						<% if cd_gestao <> "" Then
							strsql ="SELECT * From TBL_gestao where cd_codigo="&cd_gestao&""
						  	Set rs_gestao = dbconn.execute(strsql)%>
							<td colspan="2">							
							<%if not rs_gestao.EOF then%>
								<%=rs_gestao("nm_gestao")%>
							<%end if%>
						<%end if%>&nbsp;
					</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Data de entrega/retorno</b></td>
					<td colspan="2"><%=dt_retorno%></td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr bgcolor="#<%=bg_cor%>">
					<td  align="center"><b>Data de envio para Hospital</b></td>
					<td colspan="2"><%=dt_retorno_un%></td>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr><td colspan="100%"><!-- *** Inclui o registro de ocorrencias *** --><%'=cd_unidade&"."&manutencao(num_os)%>
				<%'id_origem = num_os
				id_origem = zero(cd_unidade)&"."&manutencao(num_os)
				cd_origem = 3
				pag_retorno = ".."&mem_posicao
				variaveis = "../../manutencao.asp?tipo=andamento$cd_codigo="&cd_cod_andamento&"$campo=cd_codigo"%>
				<!--#include file="../ocorrencia/ocorrencia_visualizacao.asp"-->
				</td></tr>				
				<%end if%>
				
				<%end if%>
			
			
						
			<%end if%>				
		</table>		
</BODY>
</HTML>



