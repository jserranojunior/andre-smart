<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/util.asp"-->
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<%

mes_sel = request("mes_sel")
	if mes_sel = "" Then
		mes_sel = year(now)
	end if
	
ano_sel = request("ano_sel")
	if ano_sel = "" Then
		ano_sel = year(now)
	end if
	
strcd_unidade  =  2
ordem = request("ordem")
	if ordem = "" Then
	ordem = "nm_tipo, no_patrimonio"
	end if

nm_arquivo = request("nm_arquivo")

ok = request("ok")

estilo_titulo = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:12px; color:black;"
estilo_corpo1 = "border:1px solid black; border-collapse:collapse; font-family:arial; font-size:12px;"
estilo_corpo2 = "border:0px solid white; border-collapse:collapse; font-family:arial; font-size:12px;"%>
<head>
	<title>Untitled</title>
</head>
<body>

<%Set FSO = Server.CreateObject("Scripting.FileSystemObject")
caminho_data = Server.Mappath("alimentacao/caminho_data.txt")
Set TXT1 = FSO.OpenTextFile(caminho_data)
'cria o objeto, e busca pelo TXT indicado pela variável caminho como acima


%>
	
	<table border="0" style="border-collapse:collapse;" align="center">
	<%patrimonio_array = split(txt.readALL,"$")
		For Each patrimonio_item_junto In patrimonio_array
			if patrimonio_item_junto <> "" Then
			
				if int(categoria_atual) <> int(patrimonio_cat) then
					cabeca = "0"
				end if
				
				
				patrimonio_item_len = len(patrimonio_item_junto)
					separador1 = int(instr(patrimonio_item_junto,";"))
					separador2 = int(instr(mid(patrimonio_item_junto,(separador1)+1,patrimonio_item_len),";"))
					separador3 = int(instr(mid(patrimonio_item_junto,(separador1+separador2)+1,patrimonio_item_len),";"))
					separador4 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3)+1,patrimonio_item_len),";"))
					separador5 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4)+1,patrimonio_item_len),";"))
					separador6 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4+separador5)+1,patrimonio_item_len),";"))
					separador7 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4+separador5+separador6)+1,patrimonio_item_len),";"))
					separador8 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4+separador5+separador6+separador7)+1,patrimonio_item_len),";"))
					'separador9 = int(instr(mid(patrimonio_item_junto,(separador1+separador2+separador3+separador4+separador5+separado6+separado7+separado8)+1,patrimonio_item_len),";"))
					
					patrimonio_cat =			mid(patrimonio_item_junto,1,separador1-1)
						if int(categoria_atual) <> int(patrimonio_cat) then
							cabeca = "0"
						end if
						
					
					patrimonio_unidade = 		replace(mid(patrimonio_item_junto,separador1,separador2),";","")
					patrimonio_equipamento = 	replace(mid(patrimonio_item_junto,separador1+separador2,separador3),";","")
					patrimonio_controle = 		replace(mid(patrimonio_item_junto,separador1+separador2+separador3,separador4),";","")
					patrimonio_marca = 			replace(mid(patrimonio_item_junto,separador1+separador2+separador3+separador4,separador5),";","")
					patrimonio_modelo = 		replace(mid(patrimonio_item_junto,separador1+separador2+separador3+separador4+separador5,separador6),";","")
					patrimonio_periodicidade =	replace(mid(patrimonio_item_junto,separador1+separador2+separador3+separador4+separador5+separador6,separador7),";","")
					patrimonio_mes = 			replace(mid(patrimonio_item_junto,separador1+separador2+separador3+separador4+separador5+separador6+separador7,separador8),";","")
					patrimonio_ano = 			replace(mid(patrimonio_item_junto,separador1+separador2+separador3+separador4+separador5+separador6+separador7+separador8,patrimonio_item_len),";","")
					
						if patrimonio_cat = 1 then
							patrimonio_manut = "preventiva"
						elseif patrimonio_cat = 2 then
							patrimonio_manut = "calibração"
						elseif patrimonio_cat = 3 Then
							patrimonio_manut = "segurança eletrica"
						end if%>					
						
					<%if cabeca = 0 then%>
						<!--tr><td colspan="17" colspan=>ITA20211.PQ.006 - ANEXO 3 - PLANEJAMENTO DE SERVIÇOS - VDLAP<p>OBS.: As datas marcadas são referentes a <%=patrimonio_manut%></p></td></tr-->
						<tr>		
							<td valign="bottom">
								<%if patrimonio_unidade = 2 or patrimonio_unidade = 3 or patrimonio_unidade = 15 then%><img src="../imagens/logotipo_sao_luiz.gif" alt="" width="120" height="37" border="0"><%else%>&nbsp;<%end if%></td>
							<td colspan="12" align="center" valign="bottom" style="font-family:arial; font-size:12px; color:black;"><b>PLANEJAMENTO DE SERVIÇOS REFERENTE A <%=Ucase(patrimonio_manut)%> ANO <%=patrimonio_ano%><br>VDLAP</b></td>
							<td colspan="4" align="right" valign="bottom"><img src="../imagens/logotipo_vdlap.gif" alt="" width="100" height="40" border="0"></td>
						</tr>
						<tr><td colspan="1000"><img src="../imagens/blackdot.gif" alt="" width="10" height="25" border="0"></td></tr>
						<tr>
							<!--td align="center" valign="bottom">&nbsp;</td-->
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>EQUIPAMENTO</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>CONTROLE DE<br>MANUTENÇÃO</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>MARCA</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>MODELO</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>PERIODICIDADE</b></td>		
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>JAN</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>FEV</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>MAR</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>ABR</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>MAI</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>JUN</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>JUL</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>AGO</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>SET</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>OUT</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>NOV</b></td>
							<td align="center" valign="bottom" style="<%=estilo_corpo1%>"><b>DEZ</b></td>
						</tr>
					<%cabeca = 1
					end if%>
						<%if patrimonio_controle = 0 AND patrimonio_periodicidade = 0 then
						'if patrimonio_controle = 0 AND patrimonio_modelo = 0 and patrimonio_periodicidade = 0 then%>
							<tr>
								<td style="<%=estilo_corpo1%>" colspan="5"><%=patrimonio_marca%></td>
								<td colspan="12" align="right" style="<%=estilo_corpo1%>">&nbsp;</td>
							</tr>
							<tr><td style="page_break-after:always;">&nbsp;</td></tr>
						<%else%>
							<tr>
								<td style="<%=estilo_corpo1%>">&nbsp;<%=patrimonio_equipamento%><%'=categoria_atual&"/"&patrimonio_cat&"-"&patrimonio_equipamento%></td>
								<td style="<%=estilo_corpo1%>"align="center">0<%=patrimonio_controle%></td>
								<td style="<%=estilo_corpo1%>"align="center"><%=patrimonio_marca%></td>
								<td style="<%=estilo_corpo1%>"align="center"><%=patrimonio_modelo%></td>
								<td style="<%=estilo_corpo1%>"align="center"><%=patrimonio_periodicidade%> &nbsp;<%if patrimonio_periodicidade = 1 then%>mês<%elseif patrimonio_periodicidade >= 2 then%>meses<%end if%></td>
								<%for i=1 to 12
									if int(patrimonio_mes) = i then%>
										<td style="<%=estilo_corpo1%>"><img src='../imagens/x1.gif' height=15 border='0'><%'=patrimonio_mes%></td>
									<%else%>
										<td style="<%=estilo_corpo1%>"><img src='../imagens/x0.gif' height=15 border='0'></td>										
									<%end if%>
								<%next%>
							</tr>	
						<%end if%>
						
			<%categoria_atual = patrimonio_cat
			end if
		next%>
		<tr>
							<td><img src="../imagens/blackdot.gif" alt="" width="200" height="5" border="0"></td>	
							<td><img src="../imagens/blackdot.gif" alt="" width="100" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="80" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="90" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
							<td><img src="../imagens/blackdot.gif" alt="" width="20" height="1" border="0"></td>
						</tr>
	</table>
</body>
</html>
