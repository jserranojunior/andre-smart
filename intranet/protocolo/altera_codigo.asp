<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="../includes/inc_open_connection.asp"-->
<html>
<head>
	<title>Reconfigura indice do Banco de dados</title>
</head>
<script>   
function contador() {   
  //window.setInterval('window.location.reload()', 1 * 250);   
}   
</script>   
<body onload="contador();"> 

<!--body-->
<%xsql ="SELECT  max(cd_cod_protocolo)as maximo FROM TBL_protocolo"
Set rs = dbconn.execute(xsql)
	
	sequencia = rs("maximo")
	
	if rs("maximo") > 9 Then
		sequencia = sequencia + 1
		response.write(sequencia&"//a//")
	else
		sequencia = 1
		response.write(sequencia&"//b//")
	end if%>


<%xsql ="SELECT  top 25 * FROM TBL_protocolo where cd_cod_protocolo is null order by a002_numseq,a002_numpro"
'xsql ="SELECT  * FROM TBL_protocolo where a002_numpro = 36 and a053_codung = 27 order by a002_numseq,a002_numpro"
Set rs = dbconn.execute(xsql)%>
<table>
	<tr><td>------<%=sequencia%>-------</td></tr>
	<tr>
		<td><b>Sequencia</b></td>
		<td><b>Codigo</b></td>
	</tr>
	<%while not rs.EOF
		codigo = rs("a002_numseq")
		protocolo = rs("a002_numpro")
		unidade = rs("A053_codung")
		%>
	<tr>
		<td>(***<%=sequencia%>***)</td>
		<td><%=a002_numseq%><%'=a002_numseq%></td>
	</tr>
		<%xsql_altera = "UPDATE  TBL_Protocolo SET cd_cod_protocolo = '"&sequencia&"' where a002_numseq = '"&codigo&"' and a002_numpro = '"&protocolo&"' and A053_codung = '"&unidade&"'"
		dbconn.execute(xsql_altera)
		
			'Materiais
			xsql_material = "UPDATE  TBL_Protocolo_material SET cd_protocolo = '"&sequencia&"' where A053_codung = '"&unidade&"' and a002_numseq = '"&protocolo&"'"
			dbconn.execute(xsql_material)
		
			'Patrimonio
			xsql_material = "UPDATE  TBL_Protocolo_patrimonio SET cd_protocolo = '"&sequencia&"' where cd_unidade = '"&unidade&"' and no_protocolo = '"&protocolo&"'"
			dbconn.execute(xsql_material)
		
			'Procedimento
			xsql_material = "UPDATE  TBL_Protocolo_procedimento SET cd_protocolo = '"&sequencia&"' where A053_codung = '"&unidade&"' and a002_numpro = '"&protocolo&"'"
			dbconn.execute(xsql_material)
		
		
		%>		
	<%sequencia = sequencia + 1
	rs.movenext
	wend%>
</table>




</body>
</html>
