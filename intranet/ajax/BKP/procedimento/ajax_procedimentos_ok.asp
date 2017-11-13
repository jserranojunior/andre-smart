<!--#include file="../../includes/inc_open_connection.asp"-->
<% Response.Charset="ISO-8859-1" %><!--habilita a acentuação dos ajax-->
<%'txtBusca = request("txtBusca")

procedimentos = request("procedimentos")
'procedimentos = "100"
'response.write(procedimentos&"<br>")

	'*** Apaga vírgula se for o último caractere ***
	if right(procedimentos,1) = "," Then 
	proc_len = len(procedimentos)
	proc_len = proc_len - 1
			procedimentos = Left(procedimentos,proc_len)
	end if



proc_array = split(procedimentos,",")
		For Each proc_item In proc_array
		
		if IsNumeric(proc_item) then
			xsql = "SELECT * FROM TBL_procedimento where cd_codigo="&proc_item&" order by nm_procedimento"
			Set rs_proc1 = dbconn.execute(xsql)
			while not rs_proc1.EOF
			
			cd_cod_proced = rs_proc1("cd_codigo")
			nm_procedimento = rs_proc1("nm_procedimento")%>
			
			<%=cd_cod_proced%> - <%=nm_procedimento%>&nbsp;  <!--input type="button" name="subtrair" value="-" onclick="subtrai(document.form.cd_procedimento.value,document.form.cd_procedimento_1.value,document.form.procedimentos.value)" onFocus="nextfield ='cd_procedimento';"--><br>
			<%rs_proc1.MoveNext
			 wend
			
			  rs_proc1.close
			  set rs_proc1 = nothing
		end if
		'numero = numero + 1
		
		Next

%>