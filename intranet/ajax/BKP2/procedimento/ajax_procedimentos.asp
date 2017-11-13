<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/util.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

procedimentos = request("procedimentos")

	'*** Apaga vírgula se for o último caractere ***
	if right(procedimentos,1) = "," Then 
	proc_len = len(procedimentos)
	proc_len = proc_len - 1
			procedimentos = Left(procedimentos,proc_len)
	end if	
	
	contagem = 1
	cor_num = 1%>

<table>
<tr bgcolor="#d1d1d1" style="color:white;">
					<td></td>
					<td style="width:80px;" align="right" valign="baseline">Cód.</td>
					<td style="width:300px;" valign="baseline">Descrição&nbsp;</td>
					<td bgcolor="#ffffff">&nbsp;</td>					
				</tr>
<%proc_array = split(procedimentos,",")
		For Each proc_item In proc_array
		
		if IsNumeric(proc_item) then
			xsql = "SELECT * FROM TBL_procedimento where cd_codigo="&proc_item&" order by nm_procedimento"
			Set rs_proc1 = dbconn.execute(xsql)	
			while not rs_proc1.EOF
			
			cd_cod_proced = rs_proc1("cd_codigo")
			nm_procedimento = rs_proc1("nm_procedimento")
			
			
			if cor_num = "1" then
				cor_fundo = "#ccffcc"
			else
				cor_fundo = "#ffffcc"
			end if%>
			
				<tr bgcolor="<%=cor_fundo%>">
					<td><%=zero(contagem)%></td>
					<td style="width:80px;" align="right" valign="baseline"><%=cd_cod_proced%>&nbsp;</td>
					<td style="width:300px;" valign="baseline"><%=nm_procedimento%>&nbsp;</td>
					<td bgcolor="#ffffff"><input type="button" name="subtrair" value="-" onFocus="soma('','','<%=cd_cod_proced%>',document.form.procedimentos.value)" onKeyup="nextfield ='cd_procedimento';"></td>					
				</tr>
			
			<!--input type="button" name="somar" value="-" onFocus="soma('','',document.form.cd_procedimento_2.value,document.form.procedimentos.value)" onKeyup="nextfield ='cd_procedimento';"><br-->
			<%rs_proc1.MoveNext
			 wend
			
			  rs_proc1.close
			  set rs_proc1 = nothing
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