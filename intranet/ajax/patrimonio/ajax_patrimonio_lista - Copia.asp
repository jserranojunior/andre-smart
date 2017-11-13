<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

subtotal_patrimonio = request("subtotal_patrimonio")

	'*** Apaga vírgula se for o último caractere ***
	if right(procedimentos,1) = "," Then 
	pat_len = len(procedimentos_total)
	pat_len = pat_len - 1
			procedimentos_total = Left(subtotal_patrimonio,pat_len)
	end if	
	
	contagem = 1
	cor_num = 1%>

<table>
<tr bgcolor="#d1d1d1" style="color:white;">
					<td></td>
					<td style="width:80px;" align="right" valign="baseline">Patrimônio</td>
					<td style="width:300px;" valign="baseline">Descrição&nbsp;</td>
					<td bgcolor="#ffffff">&nbsp;</td>					
				</tr>
<%pat_array = split(subtotal_patrimonio,",")
		For Each pat_item In pat_array
		
			'pat_item_len = len(pat_item)
				'if instr(pat_item,",if") = 0  then
					'pat_item = mid(pat_item,2,int(pat_item_len)-2)
				'end if
		
		if IsNumeric(pat_item) then
			xsql = "SELECT * FROM View_patrimonio_lista where cd_codigo="&pat_item&" and nm_tipo = 'o' order by nm_equipamento"
			Set rs_pat1 = dbconn.execute(xsql)	
			while not rs_pat1.EOF
			
			no_patrimonio = rs_pat1("no_patrimonio")
			nm_equipamento = rs_pat1("nm_equipamento")
			nm_tipo = rs_pat1("nm_tipo")
			
			
			if cor_num = "1" then
				cor_fundo = "#ccffcc"
			else
				cor_fundo = "#ffffcc"
			end if%>
			
				<tr bgcolor="<%=cor_fundo%>">
					<td><%=zero(contagem)%></td>
					<td style="width:80px;" align="right" valign="baseline"><%=nm_tipo&no_patrimonio%>&nbsp;<input type="hidden" name="patrimonio_excluir" value=""></td>
					<td style="width:300px;" valign="baseline"><%=nm_equipamento%>&nbsp;</td>
					<td bgcolor="#ffffff"><input type="button" name="subtrair" value="-" onFocus="soma_patrimonio('',<%=cd_patrimonio%>,document.Form.subtotal_patrimonio.value)" onKeyup="nextfield ='cd_patrimonio_busca';"><%=cd_patrimonio%></td>					
				</tr>
			
			<!--input type="button" name="somar" value="-" onFocus="soma('','',document.form.cd_procedimento_2.value,document.form.procedimentos.value)" onKeyup="nextfield ='cd_procedimento';"><br-->
			<%rs_pat1.MoveNext
			 wend
			
			  rs_pat1.close
			  set rs_pat1 = nothing
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