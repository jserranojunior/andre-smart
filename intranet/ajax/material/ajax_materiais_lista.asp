<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentua��o dos ajax-->
<%
materiais_total= request("materiais_total")

response.write(materiais_total)

	'*** Apaga v�rgula se for o �ltimo caractere ***
	if right(materiais_total,1) = "," Then 
	mat_len = len(materiais_total)
	mat_len = mat_len - 1
			materiais_total = Left(materiais_total,mat_len)
	end if
	
	contagem = 1
	cor_num = 1%>

<table>
<tr bgcolor="#d1d1d1" style="color:white;">
					<td></td>
					<td style="width:80px;" align="right" valign="baseline">C�d.</td>
					<td style="width:300px;" valign="baseline">Descri��o&nbsp;</td>
					<td style="width:50px;" align="center" valign="baseline">Qtd.</td>
					<td valign="baseline" bgcolor="#ffffff">&nbsp;</td>			
				</tr>
<%mat_array = split(materiais_total,",")
		For Each mat_item In mat_array
		
		separador = Instr(mat_item,";")
		'mat_item_qtd = replace(right(mat_item,(separador-1)),";","")
		mat_item_qtd = replace(mid(mat_item,(separador)),";","")
		
		mat_item = left(mat_item,(separador-1))
		'response.write(separador)
		
		if IsNumeric(mat_item) then
			xsql = "SELECT * FROM TBL_material where a061_codmat="&mat_item&" order by a061_nommat"
			Set rs_mat1 = dbconn.execute(xsql)	
			while not rs_mat1.EOF
			
			cd_cod_material = rs_mat1("a061_codmat")
			nm_material = rs_mat1("a061_nommat")
			qtd_material = mat_item_qtd
			
			if cor_num = "1" then
				cor_fundo = "#ccffcc"
			else
				cor_fundo = "#ffffcc"
			end if%>
			
				<tr bgcolor="<%=cor_fundo%>">
					<td><%=zero(contagem)%></td>
					<td style="width:80px;" align="right" valign="baseline"><%=cd_cod_material%>&nbsp;</td>
					<td style="width:300px;" valign="baseline"><%=nm_material%>&nbsp;</td>
					<td style="width:50px;" align="right" valign="baseline"><%=qtd_material%>&nbsp;</td>
					<td bgcolor="#ffffff"><input type="button" name="subtrair" value="-" onFocus="soma_mat('',<%=qtd_material%>,<%=cd_cod_material%>,document.getElementById('materiais_total').value)" onKeyup="nextfield ='cd_material_busca';"></td>					
				</tr>
			
			<!--input type="button" name="somar" value="-" onFocus="soma('','',document.form.cd_procedimento_2.value,document.form.procedimentos.value)" onKeyup="nextfield ='cd_procedimento';"><br-->
			<%rs_mat1.MoveNext
			 wend
			
			  rs_mat1.close
			  set rs_mat1 = nothing
		end if
		
			if cor_num = 2 then
			cor_num = 1
			else
			cor_num = cor_num + 1
			end if
			contagem = contagem + 1
		
		Next
%>
</table>			