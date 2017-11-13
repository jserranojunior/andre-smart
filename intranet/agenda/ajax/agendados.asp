<% Response.Charset="ISO-8859-1" %>
<!--#include file="../../includes/util.asp"-->
<!--#include file="../../includes/javascripts.asp"-->
<!--#include file="../../includes/inc_open_connection.asp"-->
<!--#include file="../../includes/inc_area_restrita.asp"-->
	
<script language="JavaScript">
	function esconde_agendamento() {		
			document.getElementById('ajax').style.display='none';
		}
</script>
<%cd_agenda = request("ajax")
'*** Se o valor não for vazio, processa a informação ***
if cd_agenda <> "" Then


	cd_agenda_len = len(cd_agenda)

	separador1 = instr(cd_agenda,":")
	
	cod_agenda = mid(cd_agenda,separador1+1,cd_agenda_len)


	xsql = "SELECT * FROM TBL_agenda WHERE cd_codigo = "&cod_agenda&""
	Set rs_agenda = dbconn.execute(xsql)
	'	
		if not rs_agenda.EOF Then
			cd_agenda = rs_agenda("cd_codigo")
			nm_medico = rs_agenda("nm_medico")
			nm_cirurgia = rs_agenda("nm_cirurgia")
		end if%>
	<table style="border:2px solid gray; background-color:white;" id="ajax"">
		<tr>
			<td>Med.:</td>
			<td onclick="esconde_agendamento();">
				<b><%=nm_medico%></b>
			</td>
		</tr>
		<tr>
			<td>Cir.:</td>
			<td onclick="esconde_agendamento();">
				<b><%=nm_cirurgia%></b>
			</td>
		</tr>
	</table>
<%end if%>