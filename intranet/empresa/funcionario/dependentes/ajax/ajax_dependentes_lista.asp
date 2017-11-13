<!--#include file="../../../../includes/inc_open_connection.asp"-->
<!--#include file="../../../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação do ajax-->
<%dependentes_total = request("dependentes_total")

if len(dependentes_total) > 4 then%>
<table border="1" style="border:2 solid gray; border-collapse:collapse;">
	<tr>
		<td bgcolor="#d8d8d8" style="width:30px;"><br><img src="../../imagens/px.gif" width="20" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:291px;">&nbsp;Nome<br><img src="../../imagens/px.gif" width="150" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Parentesco<br><img src="../../imagens/px.gif" width="70" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:45px;">&nbsp;Sexo<br><img src="../../imagens/px.gif" width="40" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Nascimento<br><img src="../../imagens/px.gif" width="65" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Idade<br><img src="../../imagens/px.gif" width="40" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;">&nbsp;Obs.<br><img src="../../imagens/px.gif" width="180" height="1"></td>
		<td bgcolor="#d8d8d8" style="width:15px;"><img src="../../imagens/px.gif" width="25" height="1"></td>
		
	</tr>
	
	<%'if x = 1 then
		conta_linha = 1
		dependentes_array = split(dependentes_total,"$")
		For Each dependentes_item_junto In dependentes_array
		
		if dependentes_item_junto <> "" Then
			'response.write(dependentes_item_junto)
			dependentes_item_junto = replace(dependentes_item_junto,"%20"," ")
			dependentes_item_junto = replace(dependentes_item_junto,"Ã§","ç")
			
			separador1 = int(instr(dependentes_item_junto,"!")) 'diz onde está o sinal, para separação de cada item
			separador2 = instr(dependentes_item_junto,"{")
			separador3 = instr(dependentes_item_junto,"}")
			separador4 = instr(dependentes_item_junto,"[")
			separador5 = instr(dependentes_item_junto,"]")
			separador6 = instr(dependentes_item_junto,"_")
			
			dependentes_nome =  mid(dependentes_item_junto,1,separador1) 'Mostra o primeiro compo
				dependentes_nome = replace(dependentes_nome,"!","")
			dependentes_parentesco =  mid(dependentes_item_junto,separador1,separador2-separador1)'Mostra o segundo campo
				dependentes_parentesco = replace(dependentes_parentesco,"!","")
			dependentes_nascimento = mid(dependentes_item_junto,separador2,separador3-separador2)'Mostra o terceiro campo
				dependentes_nascimento = replace(dependentes_nascimento,"{","")
					if len(dependentes_nascimento) > 4 Then
						dependentes_nascimento = zero(month(dependentes_nascimento))&"/"&zero(day(dependentes_nascimento))&"/"&year(dependentes_nascimento)
						dependentes_nascimento_mostra = zero(day(dependentes_nascimento))&"/"&zero(month(dependentes_nascimento))&"/"&year(dependentes_nascimento)
					else
						dependentes_nascimento = "null"
					end if
			dependentes_obs =   mid(dependentes_item_junto,separador3,separador4-separador3)
				dependentes_obs = replace(dependentes_obs,"}","")
			dependentes_sexo = mid(dependentes_item_junto,separador4,separador5-separador4)
				dependentes_sexo = replace(dependentes_sexo,"[","")
				%>		
		<!--tr><td colspan="100%"><%=separador1%>.<%=separador2%>.<%=separador3%>.<%=separador4%>.<%=separador5%> (Somatudo:<%=separador1+separador2%>)</td></tr-->
		
		<tr>
			<td align="right" bgcolor="#d8d8d8">&nbsp;<%=conta_linha%></td>
			<td align="left">&nbsp;<%=dependentes_nome%></td>
			<td align="left">&nbsp;
			<%strsql = "Select * From TBL_parentesco where cd_codigo='"&dependentes_parentesco&"'"'Where cd_codigo=1 or cd_codigo = 2"
				Set rs_parentesco = dbconn.execute(strsql)%>
				<%Do While Not rs_parentesco.eof
					nm_parentesco = rs_parentesco("nm_parentesco")
					cd_dependente_parentesco = rs_parentesco("cd_codigo")%>	
						<%=nm_parentesco%>
					<%rs_parentesco.movenext
				loop
						rs_parentesco.close
						Set rs_parentesco = nothing%>			
			<%'=dependentes_parentesco%></td>
			<td align="left">&nbsp;<%if dependentes_sexo = 1 then
							response.write("Masc")
						elseif dependentes_sexo = 2 then
							response.write("Fem")
						end if%>&nbsp;</td>
			<td align="center">&nbsp;<%=dependentes_nascimento_mostra%>&nbsp;</td>
			
			<%if idade <> "" Then
				idade = datediff("m",day(dependentes_nascimento)&"/"&month(dependentes_nascimento)&"/"&year(dependentes_nascimento),now)
					if idade >= 12 then
						if idade = 12 then
							tempo = "ano"
						else
							tempo = "anos"
						end if
						idade = int(idade/12)&" "&tempo
					else 
						if idade <= 1 then
							tempo = "mes"
						else
							tempo = "meses"
						end if
						idade = idade&" "&tempo
					end if
			end if%>
			<td align="center">&nbsp;<%=idade%></td>
			<td align="left">&nbsp;<%=dependentes_obs%></td>	
			<td>
				<input type="button" name="remove_dependente" value="-<%'=dependentes_item_junto%>" onclick="adiciona_dependente('','<%=dependentes_item_junto%>','','','','','','','<%=dependentes_total%>')">
			</td>
		</tr>
		
		<%
		
		total_dependentes = conta_linha
		conta_linha = conta_linha + 1
		end if%>	
	<%next
	'end if%>
		<tr><td colspan="7">Total de dependentes: &nbsp;<b><%=total_dependentes%></b></td></tr>
	
</table>
<%end if%>